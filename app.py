import os
import json
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

AR_FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")
HEADER_IMG_PATH = os.path.join(STATIC_DIR, "header.jpg")


# -------------------------
# Utilities
# -------------------------
def load_json(path: str, default):
    try:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path: str, data) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def normalize_phone(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits


def load_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # الأعمدة المطلوبة حسب صورة ملفك
    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "جهة التدريب"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة التالية غير موجودة في Excel: {', '.join(missing)}")

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()

    # صفوف ناقصة ما نبيها
    df = df[(df["رقم المتدرب"] != "") & (df["التخصص"] != "") & (df["جهة التدريب"] != "")]
    return df


def authenticate_trainee(trainee_id: str, last4: str):
    df = load_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)].head(1)
    if row.empty:
        return None, "لم يتم العثور على رقم المتدرب."

    rec = row.iloc[0].to_dict()
    phone = rec.get("رقم الجوال_norm", "")
    if len(phone) < 4 or phone[-4:] != str(last4):
        return None, "آخر 4 أرقام من الجوال غير صحيحة."

    return rec, None


# -------------------------
# Slots logic (from Excel)
# -------------------------
def build_slots_from_excel(df: pd.DataFrame):
    """
    يرجع dict بالشكل:
    {
      "تخصص1": {"جهة A": 3, "جهة B": 1, ...},
      "تخصص2": {...}
    }
    كل صف في الإكسل = فرصة واحدة لجهة تدريب ضمن تخصص.
    """
    slots = {}
    for spec, group in df.groupby("التخصص"):
        # count each entity occurrences
        counts = group["جهة التدريب"].value_counts(dropna=True).to_dict()
        # تحويل numpy types إلى int عادي
        slots[spec] = {k: int(v) for k, v in counts.items()}
    return slots


def build_used_from_assignments(assignments: dict):
    """
    assignments شكله:
    {
      "444242291": {"spec": "...", "entity": "...", "ts": "...", "name": "..."},
      ...
    }

    يرجع used:
    {
      "تخصص": {"جهة": عدد_المختارين}
    }
    """
    used = {}
    for _tid, rec in assignments.items():
        spec = (rec.get("spec") or "").strip()
        entity = (rec.get("entity") or "").strip()
        if not spec or not entity:
            continue
        used.setdefault(spec, {})
        used[spec][entity] = used[spec].get(entity, 0) + 1
    return used


def remaining_for_spec(slots_spec: dict, used_spec: dict):
    """
    slots_spec = {"جهة": total}
    used_spec  = {"جهة": used}
    يرجع قائمة خيارات للواجهة: [{"entity":..., "total":.., "used":.., "remaining":..}, ...]
    """
    options = []
    for entity, total in slots_spec.items():
        used = int(used_spec.get(entity, 0)) if used_spec else 0
        rem = int(total) - used
        if rem < 0:
            rem = 0
        options.append(
            {"entity": entity, "total": int(total), "used": used, "remaining": rem}
        )
    # ترتيب: الأكثر بقاءً أولاً
    options.sort(key=lambda x: (-x["remaining"], x["entity"]))
    return options


# -------------------------
# PDF (ReportLab)
# -------------------------
def rtl_text(s: str) -> str:
    """
    حل بسيط للـ RTL في reportlab: نعكس النص.
    (مو مثالي 100%، لكنه عملي لخطابات عربية بسيطة)
    """
    if s is None:
        return ""
    return str(s)[::-1]


def generate_letter_pdf(trainee_id: str, trainee: dict, chosen_entity: str, out_path: str):
    # تسجيل الخط
    if os.path.exists(AR_FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont("NotoNaskh", AR_FONT_PATH))
            font_name = "NotoNaskh"
        except Exception:
            font_name = "Helvetica"
    else:
        font_name = "Helvetica"

    c = canvas.Canvas(out_path, pagesize=A4)
    width, height = A4

    # Header image
    if os.path.exists(HEADER_IMG_PATH):
        # أعلى الصفحة
        c.drawImage(HEADER_IMG_PATH, 30, height - 170, width=width - 60, height=140, preserveAspectRatio=True, mask='auto')

    c.setFont(font_name, 16)
    y = height - 200

    # محتوى الخطاب (عدّله لاحقًا كما تريد)
    lines = [
        "خطاب توجيه تدريب تعاوني",
        "",
        f"اسم المتدرب: {trainee.get('اسم المتدرب','')}",
        f"رقم المتدرب: {trainee.get('رقم المتدرب','')}",
        f"التخصص: {trainee.get('التخصص','')}",
        f"جهة التدريب: {chosen_entity}",
        "",
        f"تاريخ الإصدار: {datetime.now().strftime('%Y-%m-%d')}",
    ]

    c.setFont(font_name, 14)
    for line in lines:
        # محاذاة يمين تقريبية
        txt = rtl_text(line)
        c.drawRightString(width - 50, y, txt)
        y -= 22

    c.showPage()
    c.save()


# -------------------------
# Routes
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال."
            return render_template("index.html", error=error)

        trainee, err = authenticate_trainee(trainee_id, last4)
        if err:
            return render_template("index.html", error=err)

        return redirect(url_for("dashboard", trainee_id=trainee_id))

    return render_template("index.html", error=error)


@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    trainee, err = authenticate_trainee(trainee_id, last4=request.args.get("last4", ""))
    # ملاحظة: هنا ما عندنا last4، فنجيب بيانات المتدرب من الاكسل بدون التحقق (لأن دخولك صار عبر index)
    # لو تبي حماية أعلى: خزّن session. لكن خلّينا بسيط ومستقر الآن.
    df = load_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)].head(1)
    if row.empty:
        return render_template("dashboard.html", error="لم يتم العثور على رقم المتدرب.")

    trainee = row.iloc[0].to_dict()

    assignments = load_json(ASSIGNMENTS_JSON, {})
    # هل المتدرب سبق اختار؟
    my_assignment = assignments.get(str(trainee_id))

    slots = build_slots_from_excel(df)
    used = build_used_from_assignments(assignments)

    spec = trainee.get("التخصص", "").strip()
    slots_spec = slots.get(spec, {})
    used_spec = used.get(spec, {})

    options = remaining_for_spec(slots_spec, used_spec)
    total_spec = sum(int(v) for v in slots_spec.values())
    used_spec_total = sum(int(v) for v in used_spec.values()) if used_spec else 0
    remaining_spec_total = max(total_spec - used_spec_total, 0)

    return render_template(
        "dashboard.html",
        trainee=trainee,
        options=options,
        my_assignment=my_assignment,
        total_spec=total_spec,
        remaining_spec_total=remaining_spec_total,
        error=None
    )


@app.route("/choose/<trainee_id>", methods=["POST"])
def choose(trainee_id):
    entity = (request.form.get("entity") or "").strip()
    if not entity:
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    df = load_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)].head(1)
    if row.empty:
        return render_template("dashboard.html", error="لم يتم العثور على رقم المتدرب.")

    trainee = row.iloc[0].to_dict()
    spec = trainee.get("التخصص", "").strip()

    assignments = load_json(ASSIGNMENTS_JSON, {})

    # منع إعادة الاختيار إذا اختار سابقاً (تقدر تغيرها لاحقاً)
    if str(trainee_id) in assignments:
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # تحقق من توفر الفرصة
    slots = build_slots_from_excel(df)
    used = build_used_from_assignments(assignments)

    total_for_entity = int(slots.get(spec, {}).get(entity, 0))
    used_for_entity = int(used.get(spec, {}).get(entity, 0))
    remaining = total_for_entity - used_for_entity

    if remaining <= 0:
        # ما فيه مقاعد
        return render_template(
            "dashboard.html",
            trainee=trainee,
            options=remaining_for_spec(slots.get(spec, {}), used.get(spec, {})),
            my_assignment=None,
            total_spec=sum(int(v) for v in slots.get(spec, {}).values()),
            remaining_spec_total=max(sum(int(v) for v in slots.get(spec, {}).values()) - sum(int(v) for v in used.get(spec, {}).values()), 0),
            error="عذراً، هذه الجهة انتهت فرصها. اختر جهة أخرى."
        )

    # سجّل اختيار المتدرب
    assignments[str(trainee_id)] = {
        "spec": spec,
        "entity": entity,
        "name": trainee.get("اسم المتدرب", ""),
        "ts": datetime.now().isoformat()
    }
    save_json(ASSIGNMENTS_JSON, assignments)

    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.route("/pdf/<trainee_id>", methods=["GET"])
def pdf(trainee_id):
    df = load_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)].head(1)
    if row.empty:
        return "لم يتم العثور على رقم المتدرب.", 404
    trainee = row.iloc[0].to_dict()

    assignments = load_json(ASSIGNMENTS_JSON, {})
    rec = assignments.get(str(trainee_id))
    if not rec:
        return "لا يوجد اختيار جهة لهذا المتدرب بعد.", 400

    chosen_entity = rec.get("entity", "")

    out_dir = os.path.join(BASE_DIR, "generated")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"letter_{trainee_id}.pdf")

    generate_letter_pdf(trainee_id, trainee, chosen_entity, out_path)

    return send_file(out_path, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")


if __name__ == "__main__":
    app.run(debug=True)

import os
import json
import time
import subprocess
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, send_file

from docxtpl import DocxTemplate

app = Flask(__name__)

DATA_DIR = "data"
STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
TEMPLATE_DOCX = os.path.join(DATA_DIR, "letter_template.docx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")
OUT_DIR = os.path.join(DATA_DIR, "out")

os.makedirs(OUT_DIR, exist_ok=True)


# -------------------------
# Helpers
# -------------------------
def safe_read_json(path, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def safe_write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def normalize_phone(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits


def ensure_files():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(ASSIGNMENTS_JSON):
        safe_write_json(ASSIGNMENTS_JSON, [])


def read_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError(f"ملف Excel غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "جهة التدريب"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة الأساسية غير موجودة في Excel: {', '.join(missing)}")

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم_الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)

    # optional columns
    optional = ["برنامج", "اسم المقرر", "الرقم المرجعي", "المدرب"]
    for c in optional:
        if c not in df.columns:
            df[c] = ""

    # تنظيف جهة التدريب
    df["جهة التدريب"] = df["جهة التدريب"].fillna("").astype(str).str.strip()
    df["التخصص"] = df["التخصص"].fillna("").astype(str).str.strip()

    return df


def find_trainee(df: pd.DataFrame, trainee_id: str, last4: str):
    trainee_id = (trainee_id or "").strip()
    last4 = (last4 or "").strip()

    if not trainee_id or not last4:
        return None, "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال"

    rows = df[df["رقم المتدرب"] == trainee_id]
    if rows.empty:
        return None, "لم يتم العثور على رقم المتدرب"

    row = rows.iloc[0].to_dict()
    phone = normalize_phone(row.get("رقم الجوال", ""))
    if len(phone) < 4 or phone[-4:] != last4:
        return None, "آخر 4 أرقام من الجوال غير صحيحة"

    return row, None


# -------------------------
# Slots logic (Excel = فرص)
# كل صف فيه (التخصص + جهة التدريب) = فرصة واحدة
# -------------------------
def build_inventory(df: pd.DataFrame) -> dict:
    """
    returns:
      inventory[(specialty, entity)] = total_slots (count of rows)
    """
    inv = {}
    # نعتبر فقط الصفوف التي فيها جهة تدريب غير فارغة
    df2 = df[(df["التخصص"] != "") & (df["جهة التدريب"] != "")]
    if df2.empty:
        return inv

    grouped = df2.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="count")
    for _, r in grouped.iterrows():
        inv[(str(r["التخصص"]), str(r["جهة التدريب"]))] = int(r["count"])
    return inv


def build_used_counts() -> dict:
    """
    used[(specialty, entity)] = how many trainees already assigned
    """
    used = {}
    assignments = safe_read_json(ASSIGNMENTS_JSON, [])
    for a in assignments:
        sp = (a.get("specialty") or "").strip()
        en = (a.get("entity") or "").strip()
        if not sp or not en:
            continue
        used[(sp, en)] = used.get((sp, en), 0) + 1
    return used


def remaining_for_specialty(df: pd.DataFrame, specialty: str):
    inv = build_inventory(df)
    used = build_used_counts()

    total = 0
    remaining_total = 0
    per_entity = []

    for (sp, en), cnt in inv.items():
        if sp != specialty:
            continue
        total += cnt
        u = used.get((sp, en), 0)
        rem = max(cnt - u, 0)
        remaining_total += rem
        per_entity.append({"entity": en, "remaining": rem, "total": cnt})

    # فقط الجهات التي فيها متبقي > 0
    per_entity = [x for x in per_entity if x["remaining"] > 0]
    # ترتيب: الأكثر بقايا أولاً ثم أبجدي
    per_entity.sort(key=lambda x: (-x["remaining"], x["entity"]))

    return total, remaining_total, per_entity


# -------------------------
# PDF generation
# -------------------------
def convert_docx_to_pdf(docx_path: str, out_dir: str) -> str:
    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        docx_path,
    ]
    subprocess.run(cmd, check=True)

    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("فشل تحويل PDF (لم يتم إنشاء الملف).")
    return pdf_path


def build_letter_pdf(trainee: dict, chosen_entity: str) -> str:
    if not os.path.exists(TEMPLATE_DOCX):
        raise FileNotFoundError(f"قالب الخطاب غير موجود: {TEMPLATE_DOCX}")

    # ✅ عدّل أسماء المتغيرات لتطابق placeholders في Word
    context = {
        "trainee_name": trainee.get("اسم المتدرب", ""),
        "trainee_id": trainee.get("رقم المتدرب", ""),
        "phone": normalize_phone(trainee.get("رقم الجوال", "")),
        "specialty": trainee.get("التخصص", ""),
        "program": trainee.get("برنامج", ""),
        "course_name": trainee.get("اسم المقرر", ""),
        "course_ref": trainee.get("الرقم المرجعي", ""),
        "trainer": trainee.get("المدرب", ""),
        "training_entity": chosen_entity,
        "date": datetime.now().strftime("%Y/%m/%d"),
    }

    stamp = int(time.time())
    base_name = f"letter_{trainee.get('رقم المتدرب','')}_{stamp}"
    out_docx = os.path.join(OUT_DIR, base_name + ".docx")

    doc = DocxTemplate(TEMPLATE_DOCX)
    doc.render(context)
    doc.save(out_docx)

    pdf_path = convert_docx_to_pdf(out_docx, OUT_DIR)
    return pdf_path


# -------------------------
# Routes
# -------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    ensure_files()
    error = None

    if request.method == "POST":
        try:
            df = read_students_df()
            trainee_id = request.form.get("trainee_id", "")
            last4 = request.form.get("last4", "")
            trainee, error = find_trainee(df, trainee_id, last4)
            if error:
                return render_template("index.html", error=error)
            return redirect(f"/dashboard/{trainee.get('رقم المتدرب','')}")
        except Exception as e:
            return render_template("index.html", error=str(e))

    return render_template("index.html", error=error)


@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    ensure_files()
    try:
        df = read_students_df()
        rows = df[df["رقم المتدرب"] == str(trainee_id).strip()]
        if rows.empty:
            return render_template("index.html", error="رقم المتدرب غير موجود")

        trainee = rows.iloc[0].to_dict()
        specialty = (trainee.get("التخصص") or "").strip()

        total, remaining_total, entity_options = remaining_for_specialty(df, specialty)

        if total == 0:
            # يعني لا يوجد أي صفوف فرص لهذا التخصص في Excel
            return render_template(
                "dashboard.html",
                trainee=trainee,
                total_slots=0,
                remaining_total=0,
                entity_options=[],
                error="لا توجد فرص مسجلة لهذا التخصص داخل Excel (تأكد من عمود جهة التدريب للتخصص)."
            )

        return render_template(
            "dashboard.html",
            trainee=trainee,
            total_slots=total,
            remaining_total=remaining_total,
            entity_options=entity_options,
            error=None
        )

    except Exception as e:
        return render_template("index.html", error=str(e))


@app.route("/assign", methods=["POST"])
def assign():
    ensure_files()
    try:
        df = read_students_df()

        trainee_id = (request.form.get("trainee_id") or "").strip()
        entity = (request.form.get("entity") or "").strip()
        if not trainee_id or not entity:
            return render_template("index.html", error="بيانات ناقصة")

        rows = df[df["رقم المتدرب"] == trainee_id]
        if rows.empty:
            return render_template("index.html", error="رقم المتدرب غير موجود")
        trainee = rows.iloc[0].to_dict()

        specialty = (trainee.get("التخصص") or "").strip()

        # منع تكرار الحجز لنفس المتدرب
        assignments = safe_read_json(ASSIGNMENTS_JSON, [])
        if any(a.get("trainee_id") == trainee_id for a in assignments):
            total, remaining_total, entity_options = remaining_for_specialty(df, specialty)
            return render_template(
                "dashboard.html",
                trainee=trainee,
                total_slots=total,
                remaining_total=remaining_total,
                entity_options=entity_options,
                error="تم الحجز مسبقًا لهذا المتدرب."
            )

        # تحقق أن الجهة المختارة لها متبقي
        total, remaining_total, entity_options = remaining_for_specialty(df, specialty)
        selected = next((x for x in entity_options if x["entity"] == entity), None)
        if not selected:
            return render_template(
                "dashboard.html",
                trainee=trainee,
                total_slots=total,
                remaining_total=remaining_total,
                entity_options=entity_options,
                error="هذه الجهة غير متاحة الآن (قد تكون الفرص انتهت أو ليست ضمن تخصصك)."
            )

        # سجل الحجز
        assignments.append({
            "trainee_id": trainee_id,
            "trainee_name": trainee.get("اسم المتدرب", ""),
            "specialty": specialty,
            "entity": entity,
            "ts": datetime.now().isoformat()
        })
        safe_write_json(ASSIGNMENTS_JSON, assignments)

        # طباعة PDF
        pdf_path = build_letter_pdf(trainee, entity)

        return send_file(
            pdf_path,
            as_attachment=True,
            download_name=f"خطاب_توجيه_{trainee_id}.pdf"
        )

    except subprocess.CalledProcessError:
        return render_template("index.html", error="فشل تحويل DOCX إلى PDF. تأكد من render-build.sh ووجود LibreOffice.")
    except Exception as e:
        return render_template("index.html", error=str(e))


@app.route("/admin/reset_assignments", methods=["GET"])
def reset_assignments():
    """
    تصفير الحجوزات (يعيد الفرص كما كانت)
    """
    ensure_files()
    safe_write_json(ASSIGNMENTS_JSON, [])
    return "OK: assignments reset"


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)

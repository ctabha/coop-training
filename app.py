import os
import json
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file, abort

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display


app = Flask(__name__)

DATA_DIR = "data"
STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

STATIC_DIR = "static"
HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")
AR_FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")


# -----------------------------
# Helpers
# -----------------------------
def ar(txt: str) -> str:
    """Arabic shaping + bidi for proper RTL in PDFs."""
    if txt is None:
        txt = ""
    txt = str(txt)
    reshaped = arabic_reshaper.reshape(txt)
    return get_display(reshaped)


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
        raise FileNotFoundError(f"students.xlsx not found at: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # توقع أعمدتك حسب الصورة:
    # التخصص (C) ، جهة التدريب (M) ، رقم المتدرب (H) ، اسم المتدرب (I) ، رقم الجوال (L)
    required = ["التخصص", "جهة التدريب", "رقم المتدرب", "اسم المتدرب", "رقم الجوال"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Missing columns in Excel: {', '.join(missing)}")

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()

    return df


def load_assignments() -> dict:
    # الشكل المعتمد:
    # {"trainees": {"444...": {"specialization":"...", "entity":"...", "ts":"..."}}}
    if not os.path.exists(ASSIGNMENTS_JSON):
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
            json.dump({"trainees": {}}, f, ensure_ascii=False, indent=2)

    with open(ASSIGNMENTS_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        data = {"trainees": {}}
    if "trainees" not in data or not isinstance(data["trainees"], dict):
        data["trainees"] = {}

    return data


def save_assignments(data: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def build_slots_from_excel(df: pd.DataFrame) -> dict:
    """
    الفرص تُحسب من Excel:
    لكل (تخصص + جهة تدريب) = عدد تكرارها = عدد الفرص (مقاعد)
    """
    grouped = df.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="total")
    slots = {}
    for _, row in grouped.iterrows():
        spec = row["التخصص"]
        ent = row["جهة التدريب"]
        total = int(row["total"])
        slots.setdefault(spec, {})
        slots[spec][ent] = total
    return slots


def count_used_for_spec(slots_for_spec: dict, assignments: dict, spec: str) -> dict:
    """
    يحسب كم مرة تم اختيار كل جهة داخل تخصص معيّن.
    """
    used = {ent: 0 for ent in slots_for_spec.keys()}
    for _tid, rec in assignments.get("trainees", {}).items():
        if not isinstance(rec, dict):
            continue
        if rec.get("specialization") != spec:
            continue
        ent = rec.get("entity")
        if ent in used:
            used[ent] += 1
    return used


def generate_letter_pdf(trainee: dict, chosen_entity: str, out_path: str) -> None:
    """
    PDF نظيف مع هيدر + بيانات المتدرب + الجهة المختارة.
    """
    # Register Arabic font (embed)
    if os.path.exists(AR_FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont("NotoNaskhArabic", AR_FONT_PATH))
            font_name = "NotoNaskhArabic"
        except Exception:
            font_name = "Helvetica"
    else:
        font_name = "Helvetica"

    c = canvas.Canvas(out_path, pagesize=A4)
    w, h = A4

    # Header image
    if os.path.exists(HEADER_IMG):
        c.drawImage(HEADER_IMG, 1.2 * cm, h - 5.0 * cm, width=w - 2.4 * cm, height=3.8 * cm, preserveAspectRatio=True, mask='auto')

    y = h - 6.0 * cm

    c.setFont(font_name, 18)
    c.drawCentredString(w / 2, y, ar("خطاب توجيه متدرب تدريب تعاوني"))
    y -= 1.0 * cm

    c.setFont(font_name, 12)

    # Fields
    name = trainee.get("اسم المتدرب", "")
    tid = trainee.get("رقم المتدرب", "")
    phone = trainee.get("رقم الجوال", "")
    spec = trainee.get("التخصص", "")

    # Simple table box
    box_x = 1.5 * cm
    box_w = w - 3.0 * cm
    box_h = 3.0 * cm
    c.rect(box_x, y - box_h, box_w, box_h)

    line_y = y - 0.7 * cm
    c.drawRightString(w - 2.0 * cm, line_y, ar(f"اسم المتدرب: {name}"))
    line_y -= 0.7 * cm
    c.drawRightString(w - 2.0 * cm, line_y, ar(f"رقم المتدرب: {tid}"))
    line_y -= 0.7 * cm
    c.drawRightString(w - 2.0 * cm, line_y, ar(f"رقم الجوال: {phone}"))
    line_y -= 0.7 * cm
    c.drawRightString(w - 2.0 * cm, line_y, ar(f"التخصص: {spec}"))

    y = y - box_h - 1.0 * cm

    c.setFont(font_name, 13)
    c.drawRightString(w - 2.0 * cm, y, ar(f"سعادة / {chosen_entity}"))
    y -= 0.8 * cm

    c.setFont(font_name, 12)
    body = [
        "السلام عليكم ورحمة الله وبركاته ...",
        "نفيدكم بتوجيه المتدرب الموضحة بياناته أعلاه لإتمام فترة التدريب التعاوني لدى جهتكم الموقرة.",
        "نأمل التعاون وتمكين المتدرب من تطبيق المهارات المطلوبة حسب مجال تخصصه.",
        "",
        "وتقبلوا فائق التحية ..."
    ]
    for line in body:
        c.drawRightString(w - 2.0 * cm, y, ar(line))
        y -= 0.65 * cm

    c.showPage()
    c.save()


# -----------------------------
# Routes
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال"
            return render_template("index.html", error=error)

        try:
            df = load_students_df()
        except Exception as e:
            return render_template("index.html", error=f"خطأ في ملف Excel: {e}")

        m = df[df["رقم المتدرب"] == str(trainee_id)]
        if m.empty:
            error = "لم يتم العثور على رقم المتدرب"
            return render_template("index.html", error=error)

        trainee = m.iloc[0].to_dict()
        phone = normalize_phone(trainee.get("رقم الجوال", ""))

        if len(phone) < 4 or phone[-4:] != last4:
            error = "آخر 4 أرقام من الجوال غير صحيحة"
            return render_template("index.html", error=error)

        return redirect(url_for("dashboard", trainee_id=trainee_id))

    return render_template("index.html", error=error)


@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    try:
        df = load_students_df()
    except Exception as e:
        return f"Excel error: {e}", 500

    m = df[df["رقم المتدرب"] == str(trainee_id)]
    if m.empty:
        return "Trainee not found", 404

    trainee = m.iloc[0].to_dict()
    spec = trainee.get("التخصص", "")

    slots = build_slots_from_excel(df)
    slots_for_spec = slots.get(spec, {})
    assignments = load_assignments()
    used = count_used_for_spec(slots_for_spec, assignments, spec)

    # Remaining per entity
    entities = []
    total_remaining = 0
    for ent, total in slots_for_spec.items():
        rem = int(total) - int(used.get(ent, 0))
        if rem < 0:
            rem = 0
        total_remaining += rem
        entities.append({"name": ent, "remaining": rem})

    # ترتيب (الأكثر توفر أولاً)
    entities.sort(key=lambda x: x["remaining"], reverse=True)

    current_choice = assignments["trainees"].get(str(trainee_id), {}).get("entity")

    return render_template(
        "dashboard.html",
        trainee=trainee,
        entities=entities,
        total_remaining=total_remaining,
        current_choice=current_choice
    )


@app.route("/assign/<trainee_id>", methods=["POST"])
def assign(trainee_id):
    chosen = (request.form.get("entity") or "").strip()
    if not chosen:
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # Load data
    df = load_students_df()
    m = df[df["رقم المتدرب"] == str(trainee_id)]
    if m.empty:
        abort(404)

    trainee = m.iloc[0].to_dict()
    spec = trainee.get("التخصص", "")

    slots = build_slots_from_excel(df)
    slots_for_spec = slots.get(spec, {})
    if chosen not in slots_for_spec:
        return "جهة غير متاحة لهذا التخصص", 400

    assignments = load_assignments()

    # إذا كان المتدرب اختار سابقاً: نسمح له يغير (وترجع الفرصة تلقائياً لأن العدّ مبني على assignments)
    # لكن لازم نتحقق من توفر مقعد قبل تثبيت التغيير
    used = count_used_for_spec(slots_for_spec, assignments, spec)
    remaining = int(slots_for_spec[chosen]) - int(used.get(chosen, 0))

    # لو هو نفسه مختار نفس الجهة مسبقاً → اعتبرها متاحة
    current = assignments["trainees"].get(str(trainee_id), {})
    if current.get("entity") == chosen:
        remaining = max(remaining, 1)

    if remaining <= 0:
        return "عذراً، لا توجد فرص متبقية لهذه الجهة", 400

    assignments["trainees"][str(trainee_id)] = {
        "specialization": spec,
        "entity": chosen,
        "ts": datetime.utcnow().isoformat()
    }
    save_assignments(assignments)

    # Generate PDF
    out_dir = os.path.join("tmp")
    os.makedirs(out_dir, exist_ok=True)
    pdf_path = os.path.join(out_dir, f"letter_{trainee_id}.pdf")
    generate_letter_pdf(trainee, chosen, pdf_path)

    return send_file(pdf_path, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")


@app.get("/health")
def health():
    return {"ok": True}

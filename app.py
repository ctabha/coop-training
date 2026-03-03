import os
import json
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file

from docx import Document

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.pagesizes import A4

import arabic_reshaper
from bidi.algorithm import get_display

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")
LETTER_TEMPLATE_DOCX = os.path.join(DATA_DIR, "letter_template.docx")

AR_FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")
HEADER_IMG_PATH = os.path.join(STATIC_DIR, "header.jpg")


# -----------------------------
# Helpers
# -----------------------------
def normalize_phone(x) -> str:
    """Keep only digits; handle NaN."""
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

    # أعمدة متوقعة حسب صورك:
    # "رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "برنامج", "جهة التدريب"
    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "برنامج", "جهة التدريب"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة التالية غير موجودة في Excel: {', '.join(missing)}")

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["برنامج"] = df["برنامج"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()

    # صفوف “جهة التدريب” الفاضية نشيلها
    df = df[df["جهة التدريب"].str.len() > 0].copy()

    return df


def load_assignments() -> dict:
    if not os.path.exists(ASSIGNMENTS_JSON):
        with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
        return {}
    with open(ASSIGNMENTS_JSON, "r", encoding="utf-8") as f:
        data = json.load(f)
        if not isinstance(data, dict):
            # لو الملف خربان
            return {}
        return data


def save_assignments(data: dict) -> None:
    with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def build_slots_from_excel(df: pd.DataFrame) -> dict:
    """
    يبني المقاعد من الإكسل:
    slots[التخصص][الجهة] = عدد الصفوف (المقاعد)
    """
    slots = {}
    for _, row in df.iterrows():
        spec = row["التخصص"]
        ent = row["جهة التدريب"]
        slots.setdefault(spec, {})
        slots[spec][ent] = slots[spec].get(ent, 0) + 1
    return slots


def used_counts(assignments: dict) -> dict:
    """
    used[التخصص][الجهة] = عدد من اختاروها
    """
    used = {}
    for _, a in assignments.items():
        spec = a.get("specialty")
        ent = a.get("entity")
        if not spec or not ent:
            continue
        used.setdefault(spec, {})
        used[spec][ent] = used[spec].get(ent, 0) + 1
    return used


def arabic_text(s: str) -> str:
    """Reshape + bidi for PDF drawing."""
    if not s:
        return ""
    reshaped = arabic_reshaper.reshape(s)
    return get_display(reshaped)


def fill_docx_template(out_docx_path: str, context: dict) -> None:
    """
    استبدال Placeholders مثل:
    {{TraineeName}} {{AcademicID}} {{Phone}} {{Specialty}} {{Company}} {{LetterNo}}
    """
    doc = Document(LETTER_TEMPLATE_DOCX)

    def replace_in_paragraph(paragraph):
        for k, v in context.items():
            paragraph.text = paragraph.text.replace(f"{{{{{k}}}}}", str(v))

    # paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p)

    # tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)

    doc.save(out_docx_path)


def make_pdf(out_pdf_path: str, trainee: dict, company: str, letter_no: str) -> None:
    """
    PDF بسيط وثابت:
    - يضع header.jpg أعلى الصفحة
    - يكتب بيانات المتدرب + الجهة المختارة
    """
    # Register Arabic font
    if os.path.exists(AR_FONT_PATH):
        pdfmetrics.registerFont(TTFont("NotoNaskhArabic", AR_FONT_PATH))
        font_name = "NotoNaskhArabic"
    else:
        font_name = "Helvetica"  # fallback

    c = Canvas(out_pdf_path, pagesize=A4)
    width, height = A4

    # Header image
    if os.path.exists(HEADER_IMG_PATH):
        # عرض كامل الصفحة تقريباً
        img_h = 140
        c.drawImage(HEADER_IMG_PATH, 0, height - img_h, width=width, height=img_h, preserveAspectRatio=True, mask='auto')

    y = height - 170
    c.setFont(font_name, 18)
    c.drawRightString(width - 40, y, arabic_text("خطاب توجيه متدرب تدريب تعاوني"))
    y -= 30

    c.setFont(font_name, 13)
    c.drawRightString(width - 40, y, arabic_text(f"الرقم: {letter_no}"))
    y -= 22
    c.drawRightString(width - 40, y, arabic_text(f"الاسم: {trainee.get('اسم المتدرب','')}"))
    y -= 22
    c.drawRightString(width - 40, y, arabic_text(f"رقم المتدرب: {trainee.get('رقم المتدرب','')}"))
    y -= 22
    c.drawRightString(width - 40, y, arabic_text(f"التخصص: {trainee.get('التخصص','')}"))
    y -= 22
    c.drawRightString(width - 40, y, arabic_text(f"جهة التدريب: {company}"))
    y -= 30

    c.setFont(font_name, 12)
    lines = [
        "السلام عليكم ورحمة الله وبركاته وبعد ...",
        "نفيدكم بأن المتدرب الموضح بياناته أعلاه سيتم توجيهه للتدريب التعاوني.",
        "شاكرين تعاونكم وتقبلوا فائق الاحترام.",
    ]
    for line in lines:
        c.drawRightString(width - 40, y, arabic_text(line))
        y -= 18

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
            return render_template("index.html", error=str(e))

        matches = df[df["رقم المتدرب"] == trainee_id]
        if matches.empty:
            error = "لم يتم العثور على رقم المتدرب"
            return render_template("index.html", error=error)

        # تحقق من آخر 4 أرقام من الجوال (على أول سجل للمتدرب)
        rec = matches.iloc[0].to_dict()
        phone = rec.get("رقم الجوال_norm", "")
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
        return f"خطأ في قراءة ملف الطلاب: {e}", 500

    matches = df[df["رقم المتدرب"] == str(trainee_id)]
    if matches.empty:
        return "المتدرب غير موجود", 404

    trainee = matches.iloc[0].to_dict()
    spec = trainee["التخصص"]

    slots = build_slots_from_excel(df)
    assignments = load_assignments()
    used = used_counts(assignments)

    spec_slots = slots.get(spec, {})
    spec_used = used.get(spec, {})

    # بناء قائمة الجهات للتخصص + المتبقي لكل جهة
    entities = []
    for ent_name, total in sorted(spec_slots.items(), key=lambda x: x[0]):
        u = int(spec_used.get(ent_name, 0))
        remaining = int(total) - u
        if remaining < 0:
            remaining = 0
        entities.append({"name": ent_name, "remaining": remaining})

    total_slots = sum(int(v) for v in spec_slots.values())
    remaining_slots = sum(int(item["remaining"]) for item in entities)

    return render_template(
        "dashboard.html",
        trainee=trainee,
        entities=entities,
        total_slots=total_slots,
        remaining_slots=remaining_slots,
        error=None,
        success=None
    )


@app.route("/choose/<trainee_id>", methods=["POST"])
def choose(trainee_id):
    entity = (request.form.get("entity") or "").strip()
    if not entity:
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    try:
        df = load_students_df()
    except Exception as e:
        return f"خطأ في قراءة ملف الطلاب: {e}", 500

    matches = df[df["رقم المتدرب"] == str(trainee_id)]
    if matches.empty:
        return "المتدرب غير موجود", 404

    trainee = matches.iloc[0].to_dict()
    spec = trainee["التخصص"]

    slots = build_slots_from_excel(df)
    assignments = load_assignments()
    used = used_counts(assignments)

    # تحقق أن الجهة ضمن تخصصه
    if entity not in slots.get(spec, {}):
        return render_template(
            "dashboard.html",
            trainee=trainee,
            entities=[],
            total_slots=0,
            remaining_slots=0,
            error="الجهة المختارة غير متاحة لهذا التخصص",
            success=None
        ), 400

    total = int(slots[spec][entity])
    already_used = int(used.get(spec, {}).get(entity, 0))
    remaining = total - already_used

    if remaining <= 0:
        return render_template(
            "dashboard.html",
            trainee=trainee,
            entities=[],
            total_slots=0,
            remaining_slots=0,
            error="لا توجد فرص متبقية لهذه الجهة",
            success=None
        ), 400

    # إذا سبق وسجل اختيار لهذا المتدرب، لا نكرر الإنقاص (نستبدل اختياره)
    prev = assignments.get(str(trainee_id))
    if prev and prev.get("specialty") == spec and prev.get("entity") == entity:
        # نفس الاختيار، نطبع فقط
        pass
    else:
        assignments[str(trainee_id)] = {
            "trainee_id": str(trainee_id),
            "trainee_name": trainee.get("اسم المتدرب", ""),
            "specialty": spec,
            "entity": entity,
            "created_at": datetime.utcnow().isoformat()
        }
        save_assignments(assignments)

    # توليد DOCX + PDF
    os.makedirs(os.path.join(BASE_DIR, "out"), exist_ok=True)
    out_docx = os.path.join(BASE_DIR, "out", f"letter_{trainee_id}.docx")
    out_pdf = os.path.join(BASE_DIR, "out", f"letter_{trainee_id}.pdf")

    # رقم خطاب بسيط (تقدر تعدله لاحقاً)
    letter_no = f"{trainee_id}"

    # تعبئة word (لو template موجود)
    if os.path.exists(LETTER_TEMPLATE_DOCX):
        context = {
            "TraineeName": trainee.get("اسم المتدرب", ""),
            "AcademicID": trainee.get("رقم المتدرب", ""),
            "Phone": trainee.get("رقم الجوال_norm", ""),
            "Specialty": trainee.get("التخصص", ""),
            "Company": entity,
            "LetterNo": letter_no
        }
        try:
            fill_docx_template(out_docx, context)
        except Exception:
            # لا نوقف النظام إذا فشل الـ DOCX
            pass

    # PDF (مضمون)
    make_pdf(out_pdf, trainee=trainee, company=entity, letter_no=letter_no)

    return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")


if __name__ == "__main__":
    # للتشغيل المحلي فقط
    app.run(host="0.0.0.0", port=5000, debug=True)

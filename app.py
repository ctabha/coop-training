import os
import json
import re
from datetime import datetime
from flask import Flask, render_template, request, send_file, abort
import pandas as pd

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_JSON = os.path.join(DATA_DIR, "slots.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")
FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")

app = Flask(__name__)

# ---------- Helpers ----------
def ensure_files_exist():
    missing = []
    for p in [STUDENTS_XLSX, SLOTS_JSON, ASSIGNMENTS_JSON, HEADER_IMG, FONT_PATH]:
        if not os.path.exists(p):
            missing.append(p)
    if missing:
        raise FileNotFoundError("Missing files:\n" + "\n".join(missing))

def normalize_ar(text: str) -> str:
    if text is None:
        return ""
    s = str(text).strip()
    # إزالة المسافات المتعددة
    s = re.sub(r"\s+", " ", s)
    return s

def normalize_col(col: str) -> str:
    s = normalize_ar(col)
    # توحيد بعض الاختلافات
    s = s.replace("ـ", "")
    s = s.replace("\u200f", "").replace("\u200e", "")
    s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    s = s.replace("ى", "ي")
    s = s.replace("ة", "ه")  # مفيد أحياناً
    s = s.replace(" ", "")
    return s

def load_students_df() -> pd.DataFrame:
    df = pd.read_excel(STUDENTS_XLSX, dtype=str)
    df.columns = [normalize_ar(c) for c in df.columns]

    # خريطة الأعمدة المطلوبة (نطابق حتى لو فيه اختلافات بسيطة)
    required_map = {
        "اسم المتدرب": ["اسم المتدرب", "اسم_المتدرب", "اسم"],
        "رقم المتدرب": ["رقم المتدرب", "رقم_المتدرب", "الرقم التدريبي", "الرقمالتدريبي", "الرقم"],
        "رقم الجوال": ["رقم الجوال", "رقم_الجوال", "الجوال", "رقم جوال"],
        "جهة التدريب": ["جهة التدريب", "جهه التدريب", "جهة_التدريب", "الجهة"],
        "برنامج": ["برنامج", "البرنامج"],
        "المقرر": ["المقرر", "اسم المقرر", "اسمالمقرر"],
        "اسم المقرر": ["اسم المقرر", "المقرر"],
        "الرقم المرجعي": ["الرقم المرجعي", "رقم مرجعي", "رقم_مرجعي"],
        "المدرب": ["المدرب", "اسم المدرب", "اسمالمدرب"],
        "التخصص": ["التخصص", "تخصص"],
        "القسم": ["القسم", "قسم"],
    }

    # بناء قاموس rename حسب الموجود فعلاً
    cols_norm = {c: normalize_col(c) for c in df.columns}

    rename_to = {}
    for canonical, variants in required_map.items():
        wanted = [normalize_col(v) for v in variants]
        found_col = None
        for actual, actual_n in cols_norm.items():
            if actual_n in wanted:
                found_col = actual
                break
        if found_col:
            rename_to[found_col] = canonical

    df = df.rename(columns=rename_to)

    # تحقق من أعمدة أساسية فقط (عشان الموقع يشتغل حتى لو ناقص أعمدة ثانوية)
    essential = ["اسم المتدرب", "رقم المتدرب", "رقم الجوال"]
    missing = [c for c in essential if c not in df.columns]
    if missing:
        raise ValueError(
            "الأعمدة الأساسية غير موجودة في Excel: " + " | ".join(missing) +
            "\nالأعمدة الموجودة عندك: " + " | ".join(list(df.columns))
        )

    # تنظيف
    for c in df.columns:
        df[c] = df[c].astype(str).map(normalize_ar)

    return df

def load_json(path: str, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_json(path: str, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)

def last4(phone: str) -> str:
    digits = re.sub(r"\D", "", phone or "")
    return digits[-4:] if len(digits) >= 4 else digits

def ar_text(s: str) -> str:
    """شكل عربي للـ PDF (RTL)"""
    reshaped = arabic_reshaper.reshape(s)
    return get_display(reshaped)

def pdf_register_font():
    pdfmetrics.registerFont(TTFont("NotoNaskh", FONT_PATH))

def build_pdf(output_path: str, trainee: dict, entity: str, slot_title: str):
    pdf_register_font()

    c = canvas.Canvas(output_path, pagesize=A4)
    width, height = A4

    # Header image
    if os.path.exists(HEADER_IMG):
        # نخليه أعلى الصفحة
        c.drawImage(HEADER_IMG, 30, height - 160, width - 60, 130, preserveAspectRatio=True, mask='auto')

    c.setFont("NotoNaskh", 18)
    c.drawCentredString(width / 2, height - 190, ar_text("خطاب توجيه متدرب تدريب تعاوني"))

    # Box info
    c.setLineWidth(1)
    y = height - 240

    def draw_row(label, value, x_label, x_value):
        c.setFont("NotoNaskh", 12)
        c.drawRightString(x_label, y, ar_text(label))
        c.drawRightString(x_value, y, ar_text(value))

    name = trainee.get("اسم المتدرب", "")
    tid = trainee.get("رقم المتدرب", "")
    phone = trainee.get("رقم الجوال", "")
    specialty = trainee.get("التخصص", "")
    course_ref = trainee.get("الرقم المرجعي", "")
    college_sup = trainee.get("المدرب", "")

    # جدول بسيط
    c.rect(40, y - 95, width - 80, 110)

    # سطور داخلية
    draw_row("الاسم:", name, width - 60, width - 200)
    y -= 20
    draw_row("الرقم الأكاديمي:", tid, width - 60, width - 200)
    y -= 20
    draw_row("الجوال:", phone, width - 60, width - 200)
    y -= 20
    draw_row("التخصص:", specialty, width - 60, width - 200)
    y -= 20
    draw_row("الرقم المرجعي للمقرر:", course_ref, width - 60, width - 200)
    y -= 20
    draw_row("مشرف الكلية:", college_sup, width - 60, width - 200)

    # معلومات الجهة/الفرصة
    y -= 40
    c.setFont("NotoNaskh", 14)
    c.drawRightString(width - 60, y, ar_text(f"الجهة التدريبية: {entity}"))
    y -= 25
    c.drawRightString(width - 60, y, ar_text(f"الفرصة/المسمى: {slot_title}"))

    # نص الخطاب (مختصر – عدله براحتك)
    y -= 50
    c.setFont("NotoNaskh", 13)
    lines = [
        f"سعادة/ {entity}",
        "السلام عليكم ورحمة الله وبركاته وبعد ...",
        "نفيدكم بتوجيه المتدرب الموضح بياناته أعلاه لإتمام فترة التدريب التعاوني لديكم.",
        "شاكرين تعاونكم،،",
    ]
    for line in lines:
        c.drawRightString(width - 60, y, ar_text(line))
        y -= 22

    # تاريخ
    y -= 10
    c.setFont("NotoNaskh", 12)
    c.drawRightString(width - 60, y, ar_text("التاريخ: " + datetime.now().strftime("%Y-%m-%d")))

    c.showPage()
    c.save()

def compute_remaining_slots(slots, assignments):
    # slots: list of {entity, title, capacity}
    # assignments: list of {trainee_id, entity, title}
    used = {}
    for a in assignments:
        key = (a.get("entity",""), a.get("title",""))
        used[key] = used.get(key, 0) + 1

    enriched = []
    for s in slots:
        key = (s.get("entity",""), s.get("title",""))
        cap = int(s.get("capacity", 0) or 0)
        rem = cap - used.get(key, 0)
        enriched.append({**s, "remaining": max(rem, 0)})
    return enriched

# ---------- Routes ----------
@app.route("/", methods=["GET"])
def index():
    try:
        ensure_files_exist()
        df = load_students_df()
    except Exception as e:
        return render_template("index.html", error=str(e), entities=[], slots=[])

    slots = load_json(SLOTS_JSON, default=[])
    assignments = load_json(ASSIGNMENTS_JSON, default=[])

    slots2 = compute_remaining_slots(slots, assignments)
    entities = sorted(list({s.get("entity","") for s in slots2 if s.get("entity","")}))

    return render_template("index.html", error=None, entities=entities, slots=slots2)

@app.route("/generate", methods=["POST"])
def generate():
    ensure_files_exist()
    df = load_students_df()
    slots = load_json(SLOTS_JSON, default=[])
    assignments = load_json(ASSIGNMENTS_JSON, default=[])

    trainee_id = normalize_ar(request.form.get("trainee_id", ""))
    phone4 = normalize_ar(request.form.get("phone4", ""))
    entity = normalize_ar(request.form.get("entity", ""))
    title = normalize_ar(request.form.get("title", ""))

    if not trainee_id or not phone4:
        return abort(400, "الرجاء إدخال رقم المتدرب وآخر 4 أرقام من الجوال")

    row = df[df["رقم المتدرب"].astype(str) == trainee_id]
    if row.empty:
        return abort(404, "رقم المتدرب غير موجود")

    trainee = row.iloc[0].to_dict()

    if last4(trainee.get("رقم الجوال","")) != phone4:
        return abort(403, "آخر 4 أرقام من الجوال غير صحيحة")

    # تحقق اختيار الجهة/الفرصة
    slots2 = compute_remaining_slots(slots, assignments)
    selected = None
    for s in slots2:
        if normalize_ar(s.get("entity","")) == entity and normalize_ar(s.get("title","")) == title:
            selected = s
            break
    if not selected:
        return abort(400, "الرجاء اختيار الجهة والفرصة بشكل صحيح")

    if int(selected.get("remaining", 0)) <= 0:
        return abort(400, "لا توجد شواغر متبقية لهذه الفرصة")

    # سجل الإسناد (تُنقص الفرص تلقائياً)
    assignments.append({
        "trainee_id": trainee_id,
        "entity": entity,
        "title": title,
        "ts": datetime.now().isoformat()
    })
    save_json(ASSIGNMENTS_JSON, assignments)

    # PDF output
    out_dir = os.path.join(BASE_DIR, "output")
    os.makedirs(out_dir, exist_ok=True)
    pdf_path = os.path.join(out_dir, f"letter_{trainee_id}.pdf")
    build_pdf(pdf_path, trainee, entity, title)

    return send_file(pdf_path, as_attachment=True, download_name=f"خطاب_{trainee_id}.pdf")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)

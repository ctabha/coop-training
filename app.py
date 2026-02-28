import os
import json
from io import BytesIO
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, send_file

# PDF + Arabic shaping
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_JSON = os.path.join(DATA_DIR, "slots.json")
HEADER_JPG = os.path.join(STATIC_DIR, "header.jpg")
FONT_TTF = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")


# ========= Helpers =========

def ar(text: str) -> str:
    """Arabic shaping + bidi for PDF drawing."""
    if text is None:
        text = ""
    text = str(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)

def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(os.path.join(STATIC_DIR, "fonts"), exist_ok=True)

def load_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # تنظيف أسماء الأعمدة (مهم جداً)
    df.columns = [str(c).strip() for c in df.columns]

    return df

def validate_students_columns(df: pd.DataFrame):
    required = [
        "رقم المتدرب",
        "اسم المتدرب",
        "رقم الجوال",
        "التخصص",
        "جهة التدريب",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            "الأعمدة الأساسية غير موجودة في Excel:\n"
            + " | ".join(missing)
            + "\n\nالأعمدة الموجودة عندك:\n"
            + " | ".join(df.columns.tolist())
        )

def normalize_phone_last4(v) -> str:
    s = str(v).strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits[-4:] if len(digits) >= 4 else digits

def get_or_init_slots(df: pd.DataFrame) -> dict:
    """
    slots = {
      "التخصص": {"capacity": 10, "remaining": 7},
      ...
    }
    يتولد تلقائياً لو slots.json غير موجود.
    """
    counts = df["التخصص"].astype(str).str.strip().value_counts()

    if os.path.exists(SLOTS_JSON):
        try:
            with open(SLOTS_JSON, "r", encoding="utf-8") as f:
                slots = json.load(f)
        except Exception:
            slots = {}
    else:
        slots = {}

    # تحديث/إنشاء حسب counts
    for specialty, cnt in counts.items():
        if specialty not in slots:
            slots[specialty] = {"capacity": int(cnt), "remaining": int(cnt)}
        else:
            # لو تغيّر العدد في الإكسل، نخلي capacity = الجديد
            slots[specialty]["capacity"] = int(cnt)
            # remaining لا نخليه يزيد فوق capacity
            rem = int(slots[specialty].get("remaining", cnt))
            if rem > int(cnt):
                rem = int(cnt)
            slots[specialty]["remaining"] = rem

    # حذف تخصصات لم تعد موجودة
    for sp in list(slots.keys()):
        if sp not in counts.index.astype(str).tolist():
            del slots[sp]

    with open(SLOTS_JSON, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

    return slots

def find_trainee(df: pd.DataFrame, trainee_id: str, phone_last4: str):
    tid = str(trainee_id).strip()
    last4 = str(phone_last4).strip()

    df2 = df.copy()
    df2["رقم المتدرب"] = df2["رقم المتدرب"].astype(str).str.strip()
    df2["__last4__"] = df2["رقم الجوال"].apply(normalize_phone_last4)

    row = df2[(df2["رقم المتدرب"] == tid) & (df2["__last4__"] == last4)]
    if row.empty:
        return None
    return row.iloc[0].to_dict()

def draw_pdf(trainee: dict, training_entity: str, specialty: str) -> BytesIO:
    buf = BytesIO()

    # تسجيل الخط العربي
    if os.path.exists(FONT_TTF):
        pdfmetrics.registerFont(TTFont("NotoNaskh", FONT_TTF))
        font_name = "NotoNaskh"
    else:
        font_name = "Helvetica"

    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # هيدر صورة أعلى الصفحة
    if os.path.exists(HEADER_JPG):
        # عرض الصفحة بالكامل تقريباً
        c.drawImage(HEADER_JPG, 1.2*cm, h - 6.0*cm, width=w - 2.4*cm, height=5.0*cm, preserveAspectRatio=True, mask='auto')

    c.setFont(font_name, 18)
    c.drawCentredString(w/2, h - 7.2*cm, ar("خطاب توجيه متدرب تدريب تعاوني"))

    c.setFont(font_name, 13)
    y = h - 9.0*cm

    fields = [
        ("الرقم", trainee.get("رقم المتدرب", "")),
        ("الاسم", trainee.get("اسم المتدرب", "")),
        ("التخصص", specialty),
        ("الجوال", trainee.get("رقم الجوال", "")),
        ("جهة التدريب", training_entity),
    ]

    # جدول بسيط
    x0 = 2.0*cm
    col_w = (w - 4.0*cm) / 2
    row_h = 0.9*cm

    c.setLineWidth(1)
    for i, (k, v) in enumerate(fields):
        yy = y - i*row_h
        # مربعات
        c.rect(x0, yy, col_w, row_h)
        c.rect(x0 + col_w, yy, col_w, row_h)
        c.setFont(font_name, 12)
        c.drawRightString(x0 + col_w - 0.2*cm, yy + 0.25*cm, ar(k))
        c.drawRightString(x0 + 2*col_w - 0.2*cm, yy + 0.25*cm, ar(v))

    y2 = y - len(fields)*row_h - 1.0*cm
    c.setFont(font_name, 13)
    c.drawRightString(w - 2.0*cm, y2, ar(f"سعادة/ {training_entity}"))
    y2 -= 0.8*cm
    c.drawRightString(w - 2.0*cm, y2, ar("السلام عليكم ورحمة الله وبركاته وبعد ..."))
    y2 -= 0.8*cm
    c.setFont(font_name, 12)
    c.drawRightString(w - 2.0*cm, y2, ar("نأمل منكم التكرم باستقبال المتدرب المذكور أعلاه لإكمال متطلبات التدريب التعاوني."))
    y2 -= 0.8*cm
    c.drawRightString(w - 2.0*cm, y2, ar("وتقبلوا فائق التحية والتقدير."))

    c.showPage()
    c.save()

    buf.seek(0)
    return buf


# ========= Routes =========

@app.get("/")
def home():
    return render_template("index.html", trainee=None, message=None, error=False)

@app.post("/login")
def login():
    try:
        df = load_students_df()
        validate_students_columns(df)
        slots = get_or_init_slots(df)

        trainee_id = request.form.get("trainee_id", "").strip()
        phone_last4 = request.form.get("phone_last4", "").strip()

        trainee = find_trainee(df, trainee_id, phone_last4)
        if not trainee:
            return render_template("index.html", trainee=None, message="بيانات الدخول غير صحيحة.", error=True)

        # خيارات الجهات
        entities = sorted(df["جهة التدريب"].astype(str).str.strip().unique().tolist())

        # خيارات التخصصات (الفرص) + المتبقي
        specialties = sorted(slots.keys())
        remaining_map = {sp: int(slots[sp]["remaining"]) for sp in specialties}

        return render_template(
            "index.html",
            trainee=trainee,
            entities=entities,
            specialties=specialties,
            remaining_map=remaining_map,
            message=None,
            error=False
        )
    except Exception as e:
        return render_template("index.html", trainee=None, message=str(e), error=True)

@app.post("/generate")
def generate():
    try:
        df = load_students_df()
        validate_students_columns(df)
        slots = get_or_init_slots(df)

        trainee_id = request.form.get("trainee_id", "").strip()
        # نجيب بيانات المتدرب بدون last4 هنا
        df2 = df.copy()
        df2["رقم المتدرب"] = df2["رقم المتدرب"].astype(str).str.strip()
        row = df2[df2["رقم المتدرب"] == trainee_id]
        if row.empty:
            return redirect("/")

        trainee = row.iloc[0].to_dict()

        training_entity = request.form.get("training_entity", "").strip()
        specialty = request.form.get("specialty", "").strip()

        if not training_entity:
            raise ValueError("اختر جهة التدريب.")
        if not specialty:
            raise ValueError("اختر الفرصة/التخصص.")

        if specialty not in slots:
            raise ValueError("التخصص غير موجود في slots.")

        if int(slots[specialty]["remaining"]) <= 0:
            raise ValueError(f"لا توجد مقاعد متبقية لهذا التخصص: {specialty}")

        # نقص مقعد
        slots[specialty]["remaining"] = int(slots[specialty]["remaining"]) - 1
        with open(SLOTS_JSON, "w", encoding="utf-8") as f:
            json.dump(slots, f, ensure_ascii=False, indent=2)

        pdf_buf = draw_pdf(trainee, training_entity, specialty)

        filename = f"letter_{trainee_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        return send_file(pdf_buf, as_attachment=True, download_name=filename, mimetype="application/pdf")

    except Exception as e:
        # رجّع المستخدم للصفحة الرئيسية برسالة
        return render_template("index.html", trainee=None, message=str(e), error=True)


if __name__ == "__main__":
    ensure_dirs()
    app.run(host="0.0.0.0", port=10000, debug=True)

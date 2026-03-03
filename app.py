import os
import json
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, send_file

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")
HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")


def rtl(s: str) -> str:
    """Shape + bidi for Arabic rendering in ReportLab."""
    s = "" if s is None else str(s)
    return get_display(arabic_reshaper.reshape(s))


def normalize_phone(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # لو كان رقم مخزن كـ float في الإكسل (مثال: 0500....0)
    if s.endswith(".0"):
        s = s[:-2]
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits


def load_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # تأكد من الأعمدة الأساسية حسب ملفك
    required = ["رقم المتدرب", "اسم المتدرب", "التخصص", "جهة التدريب", "رقم الجوال"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"أعمدة ناقصة في Excel: {', '.join(missing)}")

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()

    # تنظيف قيم فاضية
    df = df[df["رقم المتدرب"].notna() & (df["رقم المتدرب"] != "")]
    df = df[df["التخصص"].notna() & (df["التخصص"] != "")]
    df = df[df["جهة التدريب"].notna() & (df["جهة التدريب"] != "")]

    return df


def load_assignments() -> dict:
    if not os.path.exists(ASSIGNMENTS_JSON):
        # إنشاء افتراضي
        os.makedirs(DATA_DIR, exist_ok=True)
        with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)
        return {}

    with open(ASSIGNMENTS_JSON, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
            if isinstance(data, dict):
                return data
            return {}
        except json.JSONDecodeError:
            return {}


def save_assignments(assignments: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
        json.dump(assignments, f, ensure_ascii=False, indent=2)


def compute_slots_from_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    كل صف في Excel = مقعد.
    نرجع جدول: التخصص + جهة التدريب + total
    """
    g = df.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="total")
    return g


def compute_used_counts(assignments: dict) -> pd.DataFrame:
    """
    assignments: { trainee_id: { "spec": "...", "entity": "...", "ts": "..." } }
    نرجع جدول: spec, entity, used
    """
    rows = []
    for _, v in assignments.items():
        if isinstance(v, dict):
            spec = v.get("spec")
            ent = v.get("entity")
            if spec and ent:
                rows.append((spec, ent))
    if not rows:
        return pd.DataFrame(columns=["التخصص", "جهة التدريب", "used"])
    used = (
        pd.DataFrame(rows, columns=["التخصص", "جهة التدريب"])
        .groupby(["التخصص", "جهة التدريب"])
        .size()
        .reset_index(name="used")
    )
    return used


def make_pdf(letter_data: dict, out_path: str) -> None:
    """
    PDF عربي بسيط + الهيدر صورة + بيانات المتدرب
    """
    # تسجيل الخط مرة واحدة
    if "NotoNaskh" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("NotoNaskh", FONT_PATH))

    c = canvas.Canvas(out_path, pagesize=A4)
    w, h = A4

    # header image
    if os.path.exists(HEADER_IMG):
        # عرض كامل أعلى الصفحة
        img_w = w
        # نحافظ على نسبة تقريبية (لو صورتك طويلة/عريضة يتظبط تلقائيًا)
        c.drawImage(HEADER_IMG, 0, h - 7*cm, width=img_w, height=7*cm, preserveAspectRatio=True, mask='auto')

    c.setFont("NotoNaskh", 18)
    c.drawCentredString(w/2, h - 8.2*cm, rtl("خطاب توجيه متدرب تدريب تعاوني"))

    c.setFont("NotoNaskh", 12)

    y = h - 10*cm
    line_gap = 0.8*cm

    def right_line(label, value):
        nonlocal y
        c.drawRightString(w - 2*cm, y, rtl(f"{label}: {value}"))
        y -= line_gap

    right_line("اسم المتدرب", letter_data.get("trainee_name", ""))
    right_line("رقم المتدرب", letter_data.get("trainee_id", ""))
    right_line("التخصص", letter_data.get("spec", ""))
    right_line("جهة التدريب", letter_data.get("entity", ""))

    y -= 0.5*cm
    c.setFont("NotoNaskh", 12)
    body = [
        "سعادة/ " + str(letter_data.get("entity", "")),
        "السلام عليكم ورحمة الله وبركاته،",
        "نفيدكم بأن المتدرب الموضحة بياناته أعلاه أحد متدربي الكلية ضمن برنامج التدريب التعاوني،",
        "وعليه نأمل منكم التكرم باستقباله وتدريبه خلال فترة التدريب المحددة.",
        "شاكرين لكم تعاونكم،،"
    ]

    for s in body:
        c.drawRightString(w - 2*cm, y, rtl(s))
        y -= 0.75*cm

    y -= 0.6*cm
    c.drawRightString(w - 2*cm, y, rtl(f"التاريخ: {datetime.now().strftime('%Y-%m-%d')}"))

    c.showPage()
    c.save()


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

        match = df[df["رقم المتدرب"] == str(trainee_id)].copy()
        if match.empty:
            return render_template("index.html", error="لم يتم العثور على رقم المتدرب")

        phone = match.iloc[0]["رقم الجوال_norm"]
        if len(phone) < 4 or phone[-4:] != last4:
            return render_template("index.html", error="آخر 4 أرقام من الجوال غير صحيحة")

        return redirect(f"/dashboard/{trainee_id}")

    return render_template("index.html", error=error)


@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    try:
        df = load_students_df()
    except Exception as e:
        return f"{e}", 500

    match = df[df["رقم المتدرب"] == str(trainee_id)].copy()
    if match.empty:
        return "متدرب غير موجود", 404

    trainee_row = match.iloc[0].to_dict()
    spec = trainee_row.get("التخصص", "")

    assignments = load_assignments()
    already = None
    if str(trainee_id) in assignments and isinstance(assignments[str(trainee_id)], dict):
        already = assignments[str(trainee_id)].get("entity")

    slots_df = compute_slots_from_excel(df)
    used_df = compute_used_counts(assignments)

    merged = slots_df.merge(used_df, on=["التخصص", "جهة التدريب"], how="left")
    merged["used"] = merged["used"].fillna(0).astype(int)
    merged["remaining"] = (merged["total"].astype(int) - merged["used"]).clip(lower=0)

    # خيارات هذا التخصص فقط
    sub = merged[merged["التخصص"] == spec].copy()
    sub = sub.sort_values(["remaining", "جهة التدريب"], ascending=[False, True])

    total_for_spec = int(sub["total"].sum()) if not sub.empty else 0
    remaining_for_spec = int(sub["remaining"].sum()) if not sub.empty else 0

    options = []
    for _, r in sub.iterrows():
        options.append({
            "entity": r["جهة التدريب"],
            "total": int(r["total"]),
            "remaining": int(r["remaining"]),
        })

    return render_template(
        "dashboard.html",
        trainee=trainee_row,
        options=options,
        total_for_spec=total_for_spec,
        remaining_for_spec=remaining_for_spec,
        already_assigned=already,
        error=None
    )


@app.route("/assign/<trainee_id>", methods=["POST"])
def assign(trainee_id):
    entity = (request.form.get("entity") or "").strip()
    if not entity:
        return redirect(f"/dashboard/{trainee_id}")

    df = load_students_df()
    match = df[df["رقم المتدرب"] == str(trainee_id)].copy()
    if match.empty:
        return "متدرب غير موجود", 404

    trainee_row = match.iloc[0].to_dict()
    spec = trainee_row.get("التخصص", "")

    assignments = load_assignments()

    # لو المتدرب محجوز مسبقاً: نمنع (عشان ما ينقص مرتين)
    if str(trainee_id) in assignments:
        return redirect(f"/dashboard/{trainee_id}")

    # احسب المقاعد المتاحة الآن
    slots_df = compute_slots_from_excel(df)
    used_df = compute_used_counts(assignments)
    merged = slots_df.merge(used_df, on=["التخصص", "جهة التدريب"], how="left")
    merged["used"] = merged["used"].fillna(0).astype(int)
    merged["remaining"] = (merged["total"].astype(int) - merged["used"]).clip(lower=0)

    row = merged[(merged["التخصص"] == spec) & (merged["جهة التدريب"] == entity)]
    if row.empty:
        return redirect(f"/dashboard/{trainee_id}")

    remaining = int(row.iloc[0]["remaining"])
    if remaining <= 0:
        # لا توجد مقاعد
        return redirect(f"/dashboard/{trainee_id}")

    # سجّل الحجز
    assignments[str(trainee_id)] = {
        "spec": spec,
        "entity": entity,
        "ts": datetime.now().isoformat(timespec="seconds")
    }
    save_assignments(assignments)

    # اصنع PDF وأرسله
    tmp_dir = os.path.join(BASE_DIR, "tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    pdf_path = os.path.join(tmp_dir, f"letter_{trainee_id}.pdf")

    letter_data = {
        "trainee_name": trainee_row.get("اسم المتدرب", ""),
        "trainee_id": trainee_row.get("رقم المتدرب", ""),
        "spec": spec,
        "entity": entity,
    }
    make_pdf(letter_data, pdf_path)

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=f"خطاب_توجيه_{trainee_id}.pdf",
        mimetype="application/pdf"
    )


if __name__ == "__main__":
    app.run(debug=True)

import os
import json
import re
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, send_file

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

app = Flask(__name__)

# =========================
# Paths
# =========================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")
FONT_TTF = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")

# =========================
# Helpers
# =========================
ARABIC_DIGITS = str.maketrans("0123456789", "٠١٢٣٤٥٦٧٨٩")

def ensure_assignments_file():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(ASSIGNMENTS_JSON):
        with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)

def load_assignments():
    ensure_assignments_file()
    try:
        with open(ASSIGNMENTS_JSON, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                return data
            return {}
    except Exception:
        return {}

def save_assignments(data: dict):
    ensure_assignments_file()
    with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def normalize_phone(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = "".join(ch for ch in s if ch.isdigit())
    return s

def normalize_id(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    s = "".join(ch for ch in s if ch.isdigit())
    return s

def to_ar_digits(s: str) -> str:
    return str(s).translate(ARABIC_DIGITS)

def read_students_df():
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # الأعمدة المتوقعة حسب صورك:
    # "رقم المتدرب" ، "رقم الجوال" ، "اسم المتدرب" ، "التخصص" ، "جهة التدريب"
    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "جهة التدريب"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("الأعمدة التالية ناقصة في ملف Excel: " + ", ".join(missing))

    df["رقم_المتدرب_norm"] = df["رقم المتدرب"].apply(normalize_id)
    df["رقم_الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)

    df["التخصص"] = df["التخصص"].fillna("").astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].fillna("").astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].fillna("").astype(str).str.strip()

    # نحذف أي صف ما فيه تخصص أو ما فيه جهة تدريب
    df = df[(df["التخصص"] != "") & (df["جهة التدريب"] != "")]

    return df

def build_slots_from_excel(df: pd.DataFrame):
    """
    يبني "الفرص" من ملف Excel:
    - لكل تخصص: الجهات الموجودة في Excel تعتبر فرص.
    - عدد المقاعد لكل جهة = عدد مرات ظهورها في نفس التخصص داخل Excel.
    """
    slots = {}  # {spec: {entity: total_count}}
    for _, row in df.iterrows():
        spec = row["التخصص"]
        entity = row["جهة التدريب"]
        slots.setdefault(spec, {})
        slots[spec][entity] = slots[spec].get(entity, 0) + 1
    return slots

def used_counts(assignments: dict):
    """
    assignments format:
    {
      "444242291": {"entity": "...", "spec": "...", "ts": "..."},
      ...
    }
    """
    used = {}  # {spec: {entity: used_count}}
    for _, rec in assignments.items():
        if not isinstance(rec, dict):
            continue
        spec = (rec.get("spec") or "").strip()
        entity = (rec.get("entity") or "").strip()
        if not spec or not entity:
            continue
        used.setdefault(spec, {})
        used[spec][entity] = used[spec].get(entity, 0) + 1
    return used

def remaining_for_spec(slots_for_spec: dict, used_for_spec: dict):
    """
    returns list of tuples: (entity, remaining, total)
    """
    out = []
    used_for_spec = used_for_spec or {}
    for entity, total in slots_for_spec.items():
        u = int(used_for_spec.get(entity, 0))
        t = int(total)
        r = t - u
        if r < 0:
            r = 0
        out.append((entity, r, t))
    # ترتيب: الأكثر بقاءً أولاً ثم أبجديًا
    out.sort(key=lambda x: (-x[1], x[0]))
    return out

def register_font_once():
    # تسجيل الخط العربي مرة واحدة
    if "NotoNaskh" not in pdfmetrics.getRegisteredFontNames():
        if os.path.exists(FONT_TTF):
            pdfmetrics.registerFont(TTFont("NotoNaskh", FONT_TTF))
        else:
            # إذا الخط غير موجود لا نكسر التطبيق، لكن PDF قد لا يكون عربي مضبوط
            pass

def draw_rtl_text(c: canvas.Canvas, x_right, y, text, font_name="NotoNaskh", font_size=14):
    """
    reportlab لا يدعم RTL كامل، لكن للرسائل البسيطة نعكس النص كحل عملي.
    """
    if not text:
        return
    t = str(text)
    # قلب بسيط للنص العربي
    t = t[::-1]
    c.setFont(font_name, font_size)
    c.drawRightString(x_right, y, t)

def generate_pdf(student: dict, chosen_entity: str):
    """
    يولد PDF بخط عربي + هيدر صورة
    """
    register_font_once()

    pdf_path = os.path.join(DATA_DIR, f"letter_{student['id']}.pdf")

    c = canvas.Canvas(pdf_path, pagesize=A4)
    width, height = A4

    # هيدر
    if os.path.exists(HEADER_IMG):
        img = ImageReader(HEADER_IMG)
        # عرض كامل الصفحة تقريباً
        c.drawImage(img, 20, height - 170, width=width - 40, height=140, preserveAspectRatio=True, mask='auto')

    # نصوص الخطاب (بسيطة – تستطيع تعديلها لاحقاً)
    y = height - 200
    right = width - 40

    font = "NotoNaskh" if "NotoNaskh" in pdfmetrics.getRegisteredFontNames() else "Helvetica"

    draw_rtl_text(c, right, y, "خطاب توجيه متدرب تدريب تعاوني", font, 18)
    y -= 35

    draw_rtl_text(c, right, y, f"اسم المتدرب: {student['name']}", font, 14)
    y -= 22
    draw_rtl_text(c, right, y, f"رقم المتدرب: {student['id']}", font, 14)
    y -= 22
    draw_rtl_text(c, right, y, f"التخصص: {student['spec']}", font, 14)
    y -= 22
    draw_rtl_text(c, right, y, f"جهة التدريب: {chosen_entity}", font, 14)
    y -= 22

    today = datetime.now().strftime("%Y-%m-%d")
    draw_rtl_text(c, right, y, f"التاريخ: {to_ar_digits(today)}", font, 14)
    y -= 40

    draw_rtl_text(c, right, y, "عليه نأمل منكم التعاون وتمكين المتدرب من تنفيذ التدريب التعاوني حسب الخطة.", font, 14)
    y -= 30
    draw_rtl_text(c, right, y, "وتقبلوا تحياتنا،،", font, 14)

    c.showPage()
    c.save()

    return pdf_path

# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        trainee_id = normalize_id(request.form.get("trainee_id", ""))
        last4 = normalize_phone(request.form.get("last4", ""))

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال."
            return render_template("index.html", error=error)

        try:
            df = read_students_df()
        except Exception as e:
            return render_template("index.html", error=str(e))

        # ابحث عن المتدرب
        matches = df[df["رقم_المتدرب_norm"] == trainee_id]
        if matches.empty:
            error = "لم يتم العثور على رقم المتدرب."
            return render_template("index.html", error=error)

        # تحقق آخر 4 أرقام
        row = matches.iloc[0]
        phone = row["رقم_الجوال_norm"]
        if len(phone) < 4 or phone[-4:] != last4:
            error = "آخر 4 أرقام من الجوال غير صحيحة."
            return render_template("index.html", error=error)

        return redirect(url_for("dashboard", trainee_id=trainee_id))

    return render_template("index.html", error=error)

@app.route("/dashboard/<trainee_id>", methods=["GET", "POST"])
def dashboard(trainee_id):
    trainee_id = normalize_id(trainee_id)

    try:
        df = read_students_df()
    except Exception as e:
        return f"خطأ في قراءة ملف الطلاب: {e}", 500

    matches = df[df["رقم_المتدرب_norm"] == trainee_id]
    if matches.empty:
        return "لم يتم العثور على المتدرب.", 404

    row = matches.iloc[0]
    student = {
        "id": trainee_id,
        "name": row["اسم المتدرب"],
        "spec": row["التخصص"],
        "phone": row["رقم_الجوال_norm"],
    }

    # ابنِ الفرص من Excel
    slots = build_slots_from_excel(df)

    # لو تخصصه ما له فرص
    if student["spec"] not in slots:
        return render_template("dashboard.html",
                               student=student,
                               entities=[],
                               already=None,
                               msg="لا توجد جهات تدريب متاحة لهذا التخصص حالياً.")

    assignments = load_assignments()
    already = assignments.get(student["id"])

    # تجهيز المتبقي
    used = used_counts(assignments)
    entities_list = remaining_for_spec(slots[student["spec"]], used.get(student["spec"], {}))

    # POST: حجز + PDF
    if request.method == "POST":
        chosen = (request.form.get("entity") or "").strip()
        if not chosen:
            return render_template("dashboard.html",
                                   student=student,
                                   entities=entities_list,
                                   already=already,
                                   error="اختر جهة التدريب أولاً.")

        # لو كان محجوز مسبقاً نفس المتدرب
        if isinstance(already, dict) and already.get("entity"):
            return render_template("dashboard.html",
                                   student=student,
                                   entities=entities_list,
                                   already=already,
                                   error="تم تسجيل جهة تدريب لهذا المتدرب مسبقاً ولا يمكن التعديل.")

        # تحقق أن الجهة ضمن تخصصه ولها متبقي
        remaining_map = {e: r for (e, r, t) in entities_list}
        if chosen not in remaining_map:
            return render_template("dashboard.html",
                                   student=student,
                                   entities=entities_list,
                                   already=already,
                                   error="الجهة المختارة غير متاحة لهذا التخصص.")
        if remaining_map[chosen] <= 0:
            return render_template("dashboard.html",
                                   student=student,
                                   entities=entities_list,
                                   already=already,
                                   error="لا توجد مقاعد متبقية لهذه الجهة.")

        # سجّل الحجز
        assignments[student["id"]] = {
            "spec": student["spec"],
            "entity": chosen,
            "ts": datetime.now().isoformat(timespec="seconds")
        }
        save_assignments(assignments)

        # ولّد PDF
        try:
            pdf_path = generate_pdf(student, chosen)
        except Exception as e:
            return render_template("dashboard.html",
                                   student=student,
                                   entities=entities_list,
                                   already=assignments.get(student["id"]),
                                   error=f"تم الحجز لكن فشل توليد PDF: {e}")

        return send_file(pdf_path, as_attachment=True, download_name=f"خطاب_توجيه_{student['id']}.pdf")

    # GET
    msg = None
    if isinstance(already, dict) and already.get("entity"):
        msg = f"تم حجز جهة التدريب مسبقاً: {already.get('entity')}"

    # إجمالي/متبقي لتخصصه
    total_for_spec = sum(int(v) for v in slots[student["spec"]].values())
    used_for_spec = sum(int(v) for v in used.get(student["spec"], {}).values())
    remaining_total = max(total_for_spec - used_for_spec, 0)

    return render_template("dashboard.html",
                           student=student,
                           entities=entities_list,
                           already=already,
                           msg=msg,
                           total_for_spec=total_for_spec,
                           remaining_total=remaining_total)

if __name__ == "__main__":
    # محلي فقط
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")), debug=True)

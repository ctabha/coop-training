import os
import json
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd

from docx import Document
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ENTITIES_JSON = os.path.join(DATA_DIR, "entities.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")
LETTER_TEMPLATE_DOCX = os.path.join(DATA_DIR, "letter_template.docx")

AR_FONT_PATH = os.path.join(STATIC_DIR, "fonts", "NotoNaskhArabic-Regular.ttf")
HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret")


# -----------------------------
# Helpers: JSON (robust)
# -----------------------------
def _safe_load_json(path: str, default):
    """
    يقرأ JSON بأمان:
    - لو الملف غير موجود -> ينشأه بالقيمة الافتراضية
    - لو الملف فاضي/خربان -> يرجّع الافتراضي ويعيد كتابة الملف بشكل صحيح
    """
    os.makedirs(os.path.dirname(path), exist_ok=True)

    if not os.path.exists(path):
        _safe_write_json(path, default)
        return default

    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read().strip()
            if raw == "":
                _safe_write_json(path, default)
                return default
            return json.loads(raw)
    except Exception:
        _safe_write_json(path, default)
        return default


def _safe_write_json(path: str, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_entities():
    """
    format المطلوب:
    {
      "اسم التخصص": ["جهة 1", "جهة 2", ...],
      "تخصص آخر": [...]
    }
    """
    return _safe_load_json(ENTITIES_JSON, {})


def load_assignments():
    """
    format:
    {
      "رقم_المتدرب": {
        "spec": "...",
        "entity": "...",
        "timestamp": "ISO"
      }
    }
    """
    return _safe_load_json(ASSIGNMENTS_JSON, {})


def save_assignments(assignments: dict):
    _safe_write_json(ASSIGNMENTS_JSON, assignments)


# -----------------------------
# Helpers: Students Excel
# -----------------------------
def load_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_XLSX):
        raise FileNotFoundError("ملف data/students.xlsx غير موجود")

    df = pd.read_excel(STUDENTS_XLSX)

    # أسماء الأعمدة حسب ملفك (من صورتك)
    # F: الرقم المرجعي - G: المتدرب - H: الرقم التدريبي - I: اسم المتدرب - J: برنامج - K: رقم الجوال - L: جهة التدريب - M: (قد تكون جهة)
    # المهم عندنا: رقم المتدرب/الرقم التدريبي/الجوال/اسم/التخصص/البرنامج
    # نطبّع أسماء محتملة:
    col_map = {}
    for c in df.columns:
        s = str(c).strip()
        if s in ["الرقم التدريبي", "رقم المتدرب", "الرقم التدريبي/ رقم المتدرب"]:
            col_map["trainee_id"] = c
        elif s in ["رقم الجوال", "الجوال", "الهاتف", "رقم الهاتف"]:
            col_map["phone"] = c
        elif s in ["اسم المتدرب", "الاسم", "المتدرب"]:
            col_map["trainee_name"] = c
        elif s in ["التخصص", "التخصص/القسم"]:
            col_map["spec"] = c
        elif s in ["برنامج", "البرنامج"]:
            col_map["program"] = c

    # fallback: حاول بالترتيب الموجود عندك إن اختلفت الأسماء
    # (تقدر تعدلها لو ملفك مختلف)
    required = ["trainee_id", "phone", "trainee_name", "spec", "program"]
    missing = [k for k in required if k not in col_map]
    if missing:
        raise ValueError(f"أعمدة ناقصة في students.xlsx: {missing}. عدّل أسماء الأعمدة أو col_map في app.py")

    out = pd.DataFrame({
        "trainee_id": df[col_map["trainee_id"]].astype(str).str.strip(),
        "phone": df[col_map["phone"]].astype(str).str.strip(),
        "trainee_name": df[col_map["trainee_name"]].astype(str).str.strip(),
        "spec": df[col_map["spec"]].astype(str).str.strip(),
        "program": df[col_map["program"]].astype(str).str.strip(),
    })

    return out


def find_student(trainee_id: str, last4: str) -> dict | None:
    df = load_students_df()
    # آخر 4 أرقام من الجوال
    df["phone_last4"] = df["phone"].str.replace(" ", "").str.replace("-", "").str[-4:]
    row = df[(df["trainee_id"] == str(trainee_id).strip()) & (df["phone_last4"] == str(last4).strip())]
    if row.empty:
        return None
    r = row.iloc[0].to_dict()
    return r


# -----------------------------
# Slots logic (entity = 1 slot)
# -----------------------------
def calc_remaining_for_spec(spec: str, entities_map: dict, assignments: dict) -> dict:
    """
    يرجع dict: {entity_name: remaining_int}
    كل جهة = 1
    """
    entities = entities_map.get(spec, [])
    if not isinstance(entities, list):
        entities = []

    used = {e: 0 for e in entities}
    for tid, rec in assignments.items():
        if isinstance(rec, dict) and rec.get("spec") == spec:
            ent = rec.get("entity")
            if ent in used:
                used[ent] += 1

    remaining = {}
    for ent in entities:
        remaining[ent] = 1 - used.get(ent, 0)
    return remaining


def available_entities_for_spec(spec: str, entities_map: dict, assignments: dict) -> list:
    rem = calc_remaining_for_spec(spec, entities_map, assignments)
    return [e for e, r in rem.items() if r > 0]


# -----------------------------
# DOCX Template Text Reader
# -----------------------------
def read_docx_plain_text(docx_path: str) -> list[str]:
    """
    نقرأ نص القالب (فقرات) لاستخدامه داخل PDF.
    (نستبدل المتغيرات {{...}} لاحقاً)
    """
    if not os.path.exists(docx_path):
        return []

    doc = Document(docx_path)
    lines: list[str] = []
    for p in doc.paragraphs:
        t = (p.text or "").strip()
        if t:
            lines.append(t)

    # كذلك نصوص الجداول (لو القالب فيه جدول)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                t = (cell.text or "").strip()
                if t:
                    lines.append(t)

    # إزالة التكرار القوي
    cleaned = []
    seen = set()
    for x in lines:
        if x not in seen:
            seen.add(x)
            cleaned.append(x)

    return cleaned


# -----------------------------
# PDF Generator (Arabic)
# -----------------------------
def register_ar_font():
    if os.path.exists(AR_FONT_PATH):
        try:
            pdfmetrics.registerFont(TTFont("NotoNaskhArabic", AR_FONT_PATH))
            return True
        except Exception:
            return False
    return False


def rtl(s: str) -> str:
    """
    حل بسيط لعكس العربية في ReportLab.
    ليس تشكيل كامل، لكنه يمنع مربعات ويطلع مقروء جداً.
    """
    # لو فيه أرقام لاتعكسها بشكل غريب
    # نستخدم عكس بسيط للنص
    return s[::-1]


def generate_pdf(output_path: str, data: dict):
    """
    يولّد PDF مطابق “منطقياً” لقالبك:
    - Header image
    - Title
    - Box بيانات المتدرب
    - نص الخطاب (من letter_template.docx) بعد استبدال المتغيرات
    """
    font_ok = register_ar_font()

    c = canvas.Canvas(output_path, pagesize=A4)
    W, H = A4

    # Header image
    if os.path.exists(HEADER_IMG):
        # أعلى الصفحة
        c.drawImage(HEADER_IMG, 40, H - 170, width=W - 80, height=120, preserveAspectRatio=True, mask='auto')

    # Font
    if font_ok:
        c.setFont("NotoNaskhArabic", 18)
    else:
        c.setFont("Helvetica", 16)

    # Title
    title = "خطاب توجيه متدرب تدريب تعاوني"
    c.drawCentredString(W / 2, H - 200, rtl(title))

    # Info box
    y = H - 260
    box_x = 60
    box_w = W - 120
    box_h = 90
    c.rect(box_x, y - box_h, box_w, box_h)

    if font_ok:
        c.setFont("NotoNaskhArabic", 14)
    else:
        c.setFont("Helvetica", 12)

    lines = [
        f"اسم المتدرب: {data.get('trainee_name','')}",
        f"رقم المتدرب: {data.get('trainee_id','')}",
        f"رقم الجوال: {data.get('phone','')}",
        f"التخصص: {data.get('spec','')}",
    ]
    ty = y - 25
    for ln in lines:
        c.drawRightString(W - 80, ty, rtl(ln))
        ty -= 18

    # Letter body from DOCX
    doc_lines = read_docx_plain_text(LETTER_TEMPLATE_DOCX)

    # استبدال المتغيرات
    repl = {
        "{{trainee_name}}": data.get("trainee_name", ""),
        "{{trainee_id}}": data.get("trainee_id", ""),
        "{{phone}}": data.get("phone", ""),
        "{{training_entity}}": data.get("training_entity", ""),
        "{{course_ref}}": data.get("course_ref", ""),
        "{{college_supervisor}}": data.get("college_supervisor", ""),
        "{{start_date}}": data.get("start_date", ""),
        "{{end_date}}": data.get("end_date", ""),
    }

    def apply_repl(t: str) -> str:
        out = t
        for k, v in repl.items():
            out = out.replace(k, str(v))
        return out

    body_y = y - box_h - 35
    left_margin = 70
    right_margin = W - 70
    max_width = right_margin - left_margin

    if font_ok:
        c.setFont("NotoNaskhArabic", 13)
    else:
        c.setFont("Helvetica", 11)

    # إذا قالبك فيه نصوص كثيرة، نلفها تلقائي
    def wrap_text(text: str):
        words = text.split(" ")
        cur = ""
        out_lines = []
        for w in words:
            test = (cur + " " + w).strip()
            # قياس تقريبي
            if c.stringWidth(rtl(test), c._fontname, c._fontsize) <= max_width:
                cur = test
            else:
                if cur:
                    out_lines.append(cur)
                cur = w
        if cur:
            out_lines.append(cur)
        return out_lines

    # نختار فقط “قسم الخطاب” من القالب (تقدر تتركه كله لو تحب)
    # هنا نطبع كل الأسطر ولكن مع فلترة بسيطة للتكرارات التي من الجدول
    printed = 0
    for raw in doc_lines:
        txt = apply_repl(raw).strip()

        # تجاهل سطور المتغيرات الخام لو بقيت
        if "{{" in txt and "}}" in txt:
            continue

        for wl in wrap_text(txt):
            if body_y < 60:
                c.showPage()
                if font_ok:
                    c.setFont("NotoNaskhArabic", 13)
                else:
                    c.setFont("Helvetica", 11)
                body_y = H - 80

            c.drawRightString(right_margin, body_y, rtl(wl))
            body_y -= 18
            printed += 1

        body_y -= 8  # مسافة بين الفقرات

        if printed > 60:
            # حد أمان (عشان ما يطبع كل جدول/محتوى مكرر)
            break

    c.save()


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def index():
    return render_template("index.html")


@app.post("/")
def login():
    trainee_id = request.form.get("trainee_id", "").strip()
    last4 = request.form.get("last4", "").strip()

    if not trainee_id or not last4:
        flash("الرجاء إدخال الرقم التدريبي وآخر 4 أرقام من الجوال", "error")
        return redirect(url_for("index"))

    student = find_student(trainee_id, last4)
    if not student:
        flash("بيانات الدخول غير صحيحة", "error")
        return redirect(url_for("index"))

    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.get("/dashboard/<trainee_id>")
def dashboard(trainee_id):
    entities_map = load_entities()
    assignments = load_assignments()

    # لو JSON كان خربان سابقاً، هنا يتصلّح تلقائياً عبر safe loader
    student = None
    try:
        # نجيب الطالب بالرقم فقط (بدون last4) لأجل عرض بياناته
        df = load_students_df()
        row = df[df["trainee_id"] == str(trainee_id).strip()]
        if not row.empty:
            student = row.iloc[0].to_dict()
    except Exception:
        student = None

    if not student:
        return "المتدرب غير موجود في ملف الطلاب", 404

    spec = student["spec"]
    chosen = assignments.get(str(trainee_id))
    available = available_entities_for_spec(spec, entities_map, assignments)
    remaining_map = calc_remaining_for_spec(spec, entities_map, assignments)

    total_remaining = sum(max(0, v) for v in remaining_map.values())

    return render_template(
        "dashboard.html",
        student=student,
        chosen=chosen,
        available=available,
        remaining_map=remaining_map,
        total_remaining=total_remaining,
    )


@app.post("/choose/<trainee_id>")
def choose(trainee_id):
    entity = request.form.get("entity", "").strip()
    if not entity:
        flash("اختر جهة تدريب", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    entities_map = load_entities()
    assignments = load_assignments()

    # بيانات الطالب
    df = load_students_df()
    row = df[df["trainee_id"] == str(trainee_id).strip()]
    if row.empty:
        flash("المتدرب غير موجود", "error")
        return redirect(url_for("index"))
    student = row.iloc[0].to_dict()
    spec = student["spec"]

    # تحقق أن الجهة متاحة لهذا التخصص
    available = available_entities_for_spec(spec, entities_map, assignments)
    if entity not in available:
        flash("هذه الجهة لم تعد متاحة (تم اختيارها من متدرب آخر). اختر جهة أخرى.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # حفظ اختيار المتدرب (يستبدل القديم لو موجود)
    assignments[str(trainee_id)] = {
        "spec": spec,
        "entity": entity,
        "timestamp": datetime.utcnow().isoformat()
    }
    save_assignments(assignments)

    flash("تم حفظ اختيارك بنجاح", "ok")
    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.get("/pdf/<trainee_id>")
def pdf(trainee_id):
    assignments = load_assignments()
    rec = assignments.get(str(trainee_id))

    if not rec:
        flash("لا يوجد اختيار محفوظ لهذا المتدرب، اختر جهة أولاً.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    df = load_students_df()
    row = df[df["trainee_id"] == str(trainee_id).strip()]
    if row.empty:
        return "المتدرب غير موجود", 404
    student = row.iloc[0].to_dict()

    payload = {
        "trainee_name": student.get("trainee_name", ""),
        "trainee_id": student.get("trainee_id", ""),
        "phone": student.get("phone", ""),
        "spec": student.get("spec", ""),
        "program": student.get("program", ""),
        "training_entity": rec.get("entity", ""),
        # هذه حقول لو ودك تربطها لاحقاً من ملف/إعدادات:
        "course_ref": "",
        "college_supervisor": "",
        "start_date": "",
        "end_date": "",
    }

    out_dir = os.path.join(BASE_DIR, "tmp")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"letter_{trainee_id}.pdf")

    generate_pdf(out_path, payload)

    return send_file(out_path, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")


@app.post("/reset/<trainee_id>")
def reset_choice(trainee_id):
    assignments = load_assignments()
    if str(trainee_id) in assignments:
        del assignments[str(trainee_id)]
        save_assignments(assignments)
        flash("تم إلغاء اختيار المتدرب", "ok")
    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.get("/health")
def health():
    return {"ok": True}, 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port, debug=True)

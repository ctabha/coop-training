import os
import json
from datetime import datetime
from functools import lru_cache

import pandas as pd
from flask import (
    Flask, request, redirect, url_for,
    render_template_string, send_file, session, abort
)

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from pypdf import PdfReader, PdfWriter

import arabic_reshaper
from bidi.algorithm import get_display


# =======================
# إعدادات ومسارات
# =======================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_FILE = os.path.join(DATA_DIR, "students.xlsx")          # ملف الطلاب
TEMPLATE_PDF = os.path.join(DATA_DIR, "letter_template.pdf")     # قالب الخطاب PDF
ASSIGNMENTS_FILE = os.path.join(DATA_DIR, "assignments.json")    # حفظ اختيارات الطلاب
SLOTS_FILE = os.path.join(DATA_DIR, "slots.json")                # حفظ الفرص المتبقية

SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-please")

# مفتاح إعادة التهيئة (اختياري)
ADMIN_KEY = os.environ.get("ADMIN_KEY", "")

app = Flask(__name__)
app.secret_key = SECRET_KEY


# =======================
# أدوات مساعدة
# =======================
def ensure_dirs():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(STATIC_DIR, exist_ok=True)

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # تنظيف أسماء الأعمدة (إزالة المسافات، توحيد بعض الأسماء)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # توحيد أسماء الأعمدة المهمة (حسب ملفك students.xlsx)
    rename_map = {
        "إسم المتدرب": "اسم المتدرب",
        "اسم_المتدرب": "اسم المتدرب",
        "جهة التدريب": "جهة التدريب",
        "جهة التدريب ": "جهة التدريب",
    }
    df.rename(columns=rename_map, inplace=True)

    # تأكد أن الأعمدة الأساسية موجودة
    required = [
        "التخصص",
        "رقم المتدرب",
        "اسم المتدرب",
        "رقم الجوال",
        "جهة التدريب",
        "المدرب",
        "الرقم المرجعي",
        "اسم المقرر",
        "القسم",
        "برنامج",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"الأعمدة ناقصة في ملف students.xlsx: {missing}. الأعمدة الموجودة: {list(df.columns)}")

    # تحويل أرقام إلى نص بدون .0
    for col in ["رقم المتدرب", "رقم الجوال", "الرقم المرجعي"]:
        df[col] = df[col].apply(to_clean_str)

    # تنظيف نصوص
    for col in ["اسم المتدرب", "التخصص", "جهة التدريب", "المدرب", "اسم المقرر", "القسم", "برنامج"]:
        df[col] = df[col].astype(str).str.strip()

    return df

def to_clean_str(x) -> str:
    if pd.isna(x):
        return ""
    # إذا رقم (float/int) نحوله بدون .0
    try:
        if isinstance(x, (int,)):
            return str(x)
        if isinstance(x, float):
            if x.is_integer():
                return str(int(x))
            return str(x)
    except Exception:
        pass
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

@lru_cache(maxsize=2)
def load_students_cached(mtime: float) -> pd.DataFrame:
    df = pd.read_excel(STUDENTS_FILE)
    return normalize_columns(df)

def load_students() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_FILE):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_FILE}")
    mtime = os.path.getmtime(STUDENTS_FILE)
    return load_students_cached(mtime)

def load_json(path: str, default):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default

def save_json(path: str, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)

def init_slots_from_excel(df: pd.DataFrame) -> dict:
    """
    يبني الفرص حسب التخصص:
    كل جهة تدريب مكررة داخل نفس التخصص = عدد فرصها
    """
    slots = {}
    # group by (التخصص, جهة التدريب) count
    counts = df.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="count")
    for _, row in counts.iterrows():
        spec = row["التخصص"]
        org = row["جهة التدريب"]
        cnt = int(row["count"])
        slots.setdefault(spec, {})
        slots[spec][org] = cnt
    return slots

def get_slots(df: pd.DataFrame) -> dict:
    slots = load_json(SLOTS_FILE, default=None)
    if slots is None:
        slots = init_slots_from_excel(df)
        save_json(SLOTS_FILE, slots)
    return slots

def get_assignments() -> dict:
    return load_json(ASSIGNMENTS_FILE, default={})

def set_assignment(tid: str, payload: dict):
    assignments = get_assignments()
    assignments[tid] = payload
    save_json(ASSIGNMENTS_FILE, assignments)

def arabic_text(s: str) -> str:
    """
    تجهيز النص العربي للـ ReportLab (تشكيل + اتجاه).
    """
    if not s:
        return ""
    reshaped = arabic_reshaper.reshape(s)
    return get_display(reshaped)

def register_font():
    """
    نسجل خط يدعم العربية (غالباً DejaVuSans موجود في لينكس).
    """
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("AR", p))
            return "AR"
    # fallback
    return "Helvetica"


# =======================
# واجهات HTML (مضمنة)
# =======================
BASE_CSS = """
<style>
  body{
    font-family: Arial, sans-serif;
    background:#f7f7f7;
    margin:0;
    direction: rtl;
  }
  .top-image{
    width:100%;
    background:#fff;
    overflow:hidden;
    border-bottom:1px solid #eee;
  }
  .top-image img{
    width:100%;
    height:auto;         /* حل قص الهيدر */
    display:block;
    object-fit:contain;  /* حل قص الهيدر */
  }
  .container{
    max-width: 1100px;
    margin: 24px auto;
    background:#fff;
    padding: 28px;
    border-radius: 18px;
    box-shadow: 0 12px 30px rgba(0,0,0,.08);
    text-align:center;
  }
  h1{
    margin: 0 0 12px;
    font-size: 44px;
    font-weight: 800;
  }
  p{ color:#444; margin: 6px 0 14px; }
  .row{
    display:flex;
    gap:16px;
    justify-content:space-between;
    margin: 18px 0 8px;
    flex-wrap: wrap;
  }
  .field{
    flex: 1;
    min-width: 260px;
    text-align:right;
  }
  label{
    display:block;
    margin-bottom:8px;
    font-weight:700;
  }
  input, select{
    width:100%;
    padding: 14px 16px;
    font-size: 18px;
    border:1px solid #ddd;
    border-radius: 14px;
    outline:none;
    background:#fff;
  }
  button{
    width:100%;
    padding: 16px 18px;
    font-size: 20px;
    font-weight:700;
    border:none;
    border-radius: 16px;
    margin-top: 14px;
    cursor:pointer;
    background:#0b1630;
    color:#fff;
  }
  .error{
    color:#c00;
    margin-top: 12px;
    font-weight:700;
    white-space: pre-wrap;
  }
  .ok{
    color:#0a7a2f;
    margin-top: 12px;
    font-weight:700;
  }
  .info-box{
    background:#f3f6ff;
    border-radius: 16px;
    padding: 12px 14px;
    text-align:right;
    margin: 8px 0 14px;
  }
  .chips{
    display:flex;
    gap:10px;
    justify-content:center;
    flex-wrap:wrap;
    margin: 14px 0 10px;
  }
  .chip{
    background:#eef3ff;
    border-radius: 999px;
    padding: 10px 14px;
    font-weight:700;
  }
  a{ color:#0b3aa7; font-weight:700; }
  ul{ text-align:right; max-width: 800px; margin: 8px auto; }
</style>
"""

LOGIN_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>بوابة خطاب التوجيه - التدريب التعاوني</title>
""" + BASE_CSS + """
</head>
<body>
<div class="top-image">
  <img src="/static/header.jpg" alt="Header Image">
</div>

<div class="container">
  <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
  <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

  <form method="POST" action="/">
    <div class="row">
      <div class="field">
        <label>الرقم التدريبي</label>
        <input name="tid" placeholder="مثال: 444229747" required>
      </div>
      <div class="field">
        <label>آخر 4 أرقام من الجوال</label>
        <input name="last4" placeholder="مثال: 6101" required>
      </div>
    </div>
    <button type="submit">دخول</button>
  </form>

  {% if error %}
  <div class="error">{{error}}</div>
  {% endif %}

  <p style="margin-top:14px;color:#666;font-weight:700;">
    ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b>
  </p>
</div>
</body>
</html>
"""

CHOOSE_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>اختيار جهة التدريب</title>
""" + BASE_CSS + """
</head>
<body>
<div class="top-image">
  <img src="/static/header.jpg" alt="Header Image">
</div>

<div class="container">
  <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
  <p>اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

  <div class="chips">
    <div class="chip">المتدرب: {{name}}</div>
    <div class="chip">رقم المتدرب: {{tid}}</div>
    <div class="chip">التخصص/البرنامج: {{spec}}</div>
  </div>

  {% if already %}
    <div class="info-box">
      <div style="font-weight:800;margin-bottom:6px;">تم تسجيل اختيارك مسبقاً:</div>
      <div><b>الجهة المختارة:</b> {{already_org}}</div>
      <div style="margin-top:10px;">
        <a href="{{ url_for('letter_pdf', tid=tid) }}" target="_blank">تحميل/طباعة خطاب التوجيه PDF</a>
      </div>
    </div>
  {% endif %}

  <form method="POST" action="/choose">
    <div class="row" style="justify-content:center;">
      <div class="field" style="max-width:720px;">
        <label>جهة التدريب المتاحة</label>
        <select name="org" required {% if already %}disabled{% endif %}>
          <option value="">اختر الجهة...</option>
          {% for org, cnt in options %}
            <option value="{{org}}">{{org}} — ({{cnt}} فرصة متبقية)</option>
          {% endfor %}
        </select>
      </div>
    </div>
    <button type="submit" {% if already %}disabled{% endif %}>حفظ الاختيار</button>
  </form>

  {% if msg_ok %}
    <div class="ok">{{msg_ok}}</div>
    <div style="margin-top:10px;">
      <a href="{{ url_for('letter_pdf', tid=tid) }}" target="_blank">تحميل/طباعة خطاب التوجيه PDF</a>
    </div>
  {% endif %}

  {% if error %}
    <div class="error">{{error}}</div>
  {% endif %}

  <hr style="border:none;border-top:1px solid #eee;margin:18px 0;">

  <div class="info-box">
    <div style="font-weight:800;margin-bottom:8px;">ملخص الفرص المتبقية داخل تخصصك:</div>
    <ul>
      {% for org, cnt in summary %}
        <li>{{org}} : {{cnt}} فرصة</li>
      {% endfor %}
    </ul>
  </div>

</div>
</body>
</html>
"""


# =======================
# منطق الدخول والاختيار
# =======================
def find_student(df: pd.DataFrame, tid: str):
    row = df[df["رقم المتدرب"] == tid]
    if row.empty:
        return None
    return row.iloc[0].to_dict()

def validate_login(student: dict, last4: str) -> bool:
    phone = student.get("رقم الجوال", "")
    phone_str = str(phone).strip()
    last4_real = phone_str[-4:] if len(phone_str) >= 4 else phone_str
    return last4_real == last4

def available_options_for_specialization(slots: dict, spec: str) -> list[tuple[str, int]]:
    spec_slots = slots.get(spec, {})
    options = [(org, int(cnt)) for org, cnt in spec_slots.items() if int(cnt) > 0]
    # ترتيب تنازلي حسب عدد الفرص
    options.sort(key=lambda x: x[1], reverse=True)
    return options

def summary_for_specialization(slots: dict, spec: str) -> list[tuple[str, int]]:
    spec_slots = slots.get(spec, {})
    items = [(org, int(cnt)) for org, cnt in spec_slots.items()]
    items.sort(key=lambda x: x[1], reverse=True)
    return items


# =======================
# توليد PDF بدون soffice
# =======================
def build_overlay_pdf(out_path: str, data: dict):
    """
    ينشئ PDF شفاف فيه البيانات (Overlay) لدمجه فوق القالب.
    """
    font_name = register_font()

    c = canvas.Canvas(out_path, pagesize=A4)
    c.setFont(font_name, 12)

    # ====== إحداثيات تقريبية (يمكن تعديلها بسهولة لاحقاً)
    # الصفحة A4: (0,0) أسفل يسار
    # سنضع النص داخل خلايا الجدول تقريباً حسب قالبك.
    # إذا احتجت تعديل بسيط لاحقاً: غير الأرقام هنا فقط.

    # صفوف الجدول (y)
    y_number = 515
    y_name = 483
    y_academic = 451
    y_spec = 419
    y_phone = 387
    y_supervisor = 355
    y_ref = 323
    y_org = 291

    # عمود القيم (x) داخل الجدول جهة اليسار تقريباً
    x_value = 85

    def draw_ar(x, y, text, size=12):
        c.setFont(font_name, size)
        c.drawRightString(x, y, arabic_text(text))

    # نكتب القيم (Right aligned)
    draw_ar(500, y_number, data.get("tid", ""), 12)
    draw_ar(500, y_name, data.get("name", ""), 12)
    draw_ar(500, y_academic, data.get("tid", ""), 12)  # الرقم الأكاديمي (استخدمت رقم المتدرب)
    draw_ar(500, y_spec, data.get("spec", ""), 12)
    draw_ar(500, y_phone, data.get("phone", ""), 12)
    draw_ar(500, y_supervisor, data.get("trainer", ""), 12)
    draw_ar(500, y_ref, data.get("ref", ""), 12)
    draw_ar(500, y_org, data.get("org", ""), 12)

    # التاريخ (اختياري)
    # draw_ar(520, 250, data.get("date", ""), 11)

    c.showPage()
    c.save()

def merge_with_template(template_pdf: str, overlay_pdf: str, out_pdf: str):
    reader_template = PdfReader(template_pdf)
    reader_overlay = PdfReader(overlay_pdf)

    writer = PdfWriter()

    base_page = reader_template.pages[0]
    over_page = reader_overlay.pages[0]
    base_page.merge_page(over_page)
    writer.add_page(base_page)

    with open(out_pdf, "wb") as f:
        writer.write(f)


# =======================
# Routes
# =======================
@app.route("/", methods=["GET", "POST"])
def login():
    try:
        df = load_students()
    except Exception as e:
        return render_template_string(LOGIN_PAGE, error=f"حدث خطأ أثناء التحميل/التحقق:\n{e}")

    if request.method == "GET":
        return render_template_string(LOGIN_PAGE, error=None)

    tid = to_clean_str(request.form.get("tid", ""))
    last4 = to_clean_str(request.form.get("last4", ""))

    if not tid or not last4:
        return render_template_string(LOGIN_PAGE, error="الرجاء إدخال الرقم التدريبي وآخر 4 أرقام من الجوال.")

    student = find_student(df, tid)
    if not student:
        return render_template_string(LOGIN_PAGE, error="لم يتم العثور على الرقم التدريبي في ملف الطلاب.")

    if not validate_login(student, last4):
        return render_template_string(LOGIN_PAGE, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

    # نجاح: خزّن tid في session
    session["tid"] = tid
    return redirect(url_for("choose"))


@app.route("/choose", methods=["GET", "POST"])
def choose():
    tid = session.get("tid")
    if not tid:
        return redirect(url_for("login"))

    try:
        df = load_students()
        slots = get_slots(df)
        assignments = get_assignments()
    except Exception as e:
        return render_template_string(CHOOSE_PAGE, error=f"خطأ أثناء التحميل:\n{e}", name="", tid=tid, spec="", options=[], summary=[], already=False)

    student = find_student(df, tid)
    if not student:
        session.pop("tid", None)
        return redirect(url_for("login"))

    name = student.get("اسم المتدرب", "")
    spec = student.get("التخصص", "") or student.get("برنامج", "")
    phone = student.get("رقم الجوال", "")
    trainer = student.get("المدرب", "")
    ref = student.get("الرقم المرجعي", "")

    already = tid in assignments
    already_org = assignments.get(tid, {}).get("org", "")

    if request.method == "GET":
        options = available_options_for_specialization(slots, spec)
        summary = summary_for_specialization(slots, spec)
        return render_template_string(
            CHOOSE_PAGE,
            name=name, tid=tid, spec=spec,
            options=options,
            summary=summary,
            already=already,
            already_org=already_org,
            msg_ok=None,
            error=None
        )

    # POST: حفظ الاختيار (إذا لم يكن مختار سابقاً)
    if already:
        return redirect(url_for("choose"))

    org = (request.form.get("org") or "").strip()
    if not org:
        options = available_options_for_specialization(slots, spec)
        summary = summary_for_specialization(slots, spec)
        return render_template_string(
            CHOOSE_PAGE,
            name=name, tid=tid, spec=spec,
            options=options, summary=summary,
            already=False, already_org="",
            msg_ok=None,
            error="الرجاء اختيار جهة تدريب."
        )

    # تحقق أن الفرصة ما زالت متاحة
    if spec not in slots or org not in slots[spec] or int(slots[spec][org]) <= 0:
        options = available_options_for_specialization(slots, spec)
        summary = summary_for_specialization(slots, spec)
        return render_template_string(
            CHOOSE_PAGE,
            name=name, tid=tid, spec=spec,
            options=options, summary=summary,
            already=False, already_org="",
            msg_ok=None,
            error="هذه الجهة لم تعد متاحة الآن. اختر جهة أخرى."
        )

    # خصم فرصة + حفظ
    slots[spec][org] = int(slots[spec][org]) - 1
    save_json(SLOTS_FILE, slots)

    payload = {
        "tid": tid,
        "name": name,
        "spec": spec,
        "phone": phone,
        "trainer": trainer,
        "ref": ref,
        "org": org,
        "ts": datetime.utcnow().isoformat() + "Z"
    }
    set_assignment(tid, payload)

    options = available_options_for_specialization(slots, spec)
    summary = summary_for_specialization(slots, spec)
    return render_template_string(
        CHOOSE_PAGE,
        name=name, tid=tid, spec=spec,
        options=options,
        summary=summary,
        already=True,
        already_org=org,
        msg_ok="تم حفظ اختيارك بنجاح.",
        error=None
    )


@app.route("/letter/<tid>", methods=["GET"])
def letter_pdf(tid: str):
    # تأمين بسيط: لازم يكون نفس المستخدم داخل الجلسة
    if session.get("tid") != tid:
        abort(403)

    if not os.path.exists(TEMPLATE_PDF):
        abort(500, description=f"قالب PDF غير موجود: {TEMPLATE_PDF}. ضع القالب داخل data باسم letter_template.pdf")

    assignments = get_assignments()
    if tid not in assignments:
        abort(404)

    a = assignments[tid]

    # توليد overlay ثم دمجه مع القالب
    tmp_overlay = os.path.join(DATA_DIR, f"_overlay_{tid}.pdf")
    out_pdf = os.path.join(DATA_DIR, f"letter_{tid}.pdf")

    build_overlay_pdf(tmp_overlay, {
        "tid": a.get("tid", ""),
        "name": a.get("name", ""),
        "spec": a.get("spec", ""),
        "phone": a.get("phone", ""),
        "trainer": a.get("trainer", ""),
        "ref": a.get("ref", ""),
        "org": a.get("org", ""),
        "date": datetime.now().strftime("%Y/%m/%d")
    })
    merge_with_template(TEMPLATE_PDF, tmp_overlay, out_pdf)

    # تنظيف overlay
    try:
        os.remove(tmp_overlay)
    except Exception:
        pass

    return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{tid}.pdf")


@app.route("/logout")
def logout():
    session.pop("tid", None)
    return redirect(url_for("login"))


@app.route("/admin/reset")
def admin_reset():
    """
    إعادة تهيئة الفرص والاختيارات.
    استخدم: /admin/reset?key=YOUR_ADMIN_KEY
    """
    if not ADMIN_KEY:
        abort(403)
    key = request.args.get("key", "")
    if key != ADMIN_KEY:
        abort(403)

    df = load_students()
    slots = init_slots_from_excel(df)
    save_json(SLOTS_FILE, slots)
    save_json(ASSIGNMENTS_FILE, {})
    return "OK: reset done"


if __name__ == "__main__":
    ensure_dirs()
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

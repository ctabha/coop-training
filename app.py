import os
import json
from io import BytesIO
from datetime import datetime

import pandas as pd
from flask import (
    Flask, request, redirect, url_for,
    render_template_string, send_file
)

# PDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Arabic shaping (for proper Arabic in PDF)
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    AR_SUPPORT = True
except Exception:
    AR_SUPPORT = False


# ---------------------------
# Config
# ---------------------------
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_FILE = os.path.join(DATA_DIR, "students.xlsx")  # ✅ important
ASSIGNMENTS_FILE = os.path.join(DATA_DIR, "assignments.json")

HEADER_IMG = os.path.join(STATIC_DIR, "header.jpg")

REQUIRED_COLS_CANONICAL = [
    "رقم المتدرب",
    "إسم المتدرب",
    "رقم الجوال",
    "التخصص",
    "برنامج",
    "المدرب",
    "الرقم المرجعي",
    "اسم المقرر",
    "جهة التدريب",
]


# ---------------------------
# Helpers
# ---------------------------
def ensure_data_files():
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(STATIC_DIR, exist_ok=True)
    if not os.path.exists(ASSIGNMENTS_FILE):
        with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)


def norm_col(col: str) -> str:
    # normalize Arabic variants and trim spaces
    c = str(col).strip()
    c = c.replace("اسم المتدرب", "إسم المتدرب")
    c = c.replace("جهة التدريب ", "جهة التدريب")  # remove trailing space variant
    c = c.replace("جهة التدريب  ", "جهة التدريب")
    return c


def safe_str(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    # if Excel number came like '444242291.0'
    if s.endswith(".0"):
        s = s[:-2]
    return s.strip()


def load_students_df() -> pd.DataFrame:
    if not os.path.exists(STUDENTS_FILE):
        raise FileNotFoundError(f"الملف غير موجود: {STUDENTS_FILE}")

    # Read everything as string to preserve leading zeros in phone if any
    df = pd.read_excel(STUDENTS_FILE, dtype=str)
    df.columns = [norm_col(c) for c in df.columns]

    # also normalize possible trailing spaces inside column name
    if "جهة التدريب " in df.columns and "جهة التدريب" not in df.columns:
        df = df.rename(columns={"جهة التدريب ": "جهة التدريب"})

    # strip cells
    for c in df.columns:
        df[c] = df[c].map(safe_str)

    return df


def validate_columns(df: pd.DataFrame):
    # We only require the columns we truly use
    needed = {
        "رقم المتدرب",
        "إسم المتدرب",
        "رقم الجوال",
        "التخصص",
        "برنامج",
        "جهة التدريب",
        "المدرب",
        "الرقم المرجعي",
        "اسم المقرر",
    }
    missing = [c for c in needed if c not in df.columns]
    if missing:
        raise KeyError(
            f"لم أجد الأعمدة المطلوبة: {missing}. الأعمدة الموجودة: {df.columns.tolist()}"
        )


def last4_of_phone(phone: str) -> str:
    p = safe_str(phone)
    digits = "".join(ch for ch in p if ch.isdigit())
    return digits[-4:] if len(digits) >= 4 else digits


def load_assignments() -> dict:
    ensure_data_files()
    try:
        with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_assignments(data: dict):
    ensure_data_files()
    with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def capacities_by_specialization(df: pd.DataFrame) -> dict:
    """
    الفرص تُحسب من ملف الإكسل:
    كل (تخصص + جهة تدريب) مكرر => يزيد عدد الفرص.
    """
    df2 = df.copy()
    df2["جهة التدريب"] = df2["جهة التدريب"].map(safe_str)
    df2["التخصص"] = df2["التخصص"].map(safe_str)

    df2 = df2[(df2["جهة التدريب"] != "") & (df2["التخصص"] != "")]
    grp = df2.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="capacity")

    result = {}
    for _, row in grp.iterrows():
        spec = row["التخصص"]
        org = row["جهة التدريب"]
        cap = int(row["capacity"])
        result.setdefault(spec, {})[org] = cap
    return result


def remaining_for_student(df: pd.DataFrame, student: dict) -> tuple[list, list]:
    """
    returns:
    - available_orgs: list of (org, remaining, capacity)
    - summary: list of (org, remaining)
    """
    caps = capacities_by_specialization(df)
    spec = student.get("التخصص", "")
    org_caps = caps.get(spec, {})

    assignments = load_assignments()

    # count assignments per spec+org
    used = {}
    for tid, info in assignments.items():
        s = info.get("التخصص", "")
        o = info.get("جهة التدريب", "")
        if s and o:
            used.setdefault(s, {}).setdefault(o, 0)
            used[s][o] += 1

    available = []
    for org, cap in org_caps.items():
        used_count = used.get(spec, {}).get(org, 0)
        rem = max(0, cap - used_count)
        if rem > 0:
            available.append((org, rem, cap))

    available.sort(key=lambda x: (-x[1], x[0]))
    summary = [(org, rem) for (org, rem, _cap) in available]
    return available, summary


def arabic(s: str) -> str:
    s = safe_str(s)
    if not s:
        return ""
    if not AR_SUPPORT:
        return s
    reshaped = arabic_reshaper.reshape(s)
    return get_display(reshaped)


def register_font():
    """
    Try to register DejaVuSans (usually exists on Linux).
    If not found, PDF will still generate but Arabic might not render well.
    """
    candidates = [
        os.path.join(BASE_DIR, "DejaVuSans.ttf"),
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            try:
                pdfmetrics.registerFont(TTFont("DejaVu", p))
                return "DejaVu"
            except Exception:
                pass
    return "Helvetica"


def make_letter_pdf(student: dict, chosen_org: str) -> bytes:
    font_name = register_font()

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # Header image (fit width, keep aspect, prevent cropping)
    y = height - 1.2 * cm
    if os.path.exists(HEADER_IMG):
        try:
            # draw image full width with fixed height area
            img_h = 4.0 * cm
            c.drawImage(HEADER_IMG, 0, height - img_h, width=width, height=img_h, preserveAspectRatio=True, anchor='n')
            y = height - img_h - 1.0 * cm
        except Exception:
            y = height - 2.0 * cm

    c.setFont(font_name, 18)
    c.drawCentredString(width / 2, y, arabic("خطاب توجيه متدرب تدريب تعاوني"))
    y -= 1.0 * cm

    c.setFont(font_name, 12)

    # Table-like info
    rows = [
        ("الرقم", student.get("رقم المتدرب", "")),
        ("الاسم", student.get("إسم المتدرب", "")),
        ("الرقم الأكاديمي", student.get("رقم المتدرب", "")),
        ("التخصص", student.get("التخصص", "")),
        ("جوال", student.get("رقم الجوال", "")),
        ("اسم المشرف من الكلية", student.get("المدرب", "")),
        ("الرقم المرجعي للمقرر", student.get("الرقم المرجعي", "")),
        ("جهة التدريب المختارة", chosen_org),
    ]

    x0 = 2.0 * cm
    col1 = 6.0 * cm
    col2 = width - 2.0 * cm

    box_h = 0.9 * cm
    for label, value in rows:
        c.rect(x0, y - box_h, col2 - x0, box_h, stroke=1, fill=0)
        c.setFont(font_name, 11)
        c.drawRightString(col2 - 0.3 * cm, y - 0.62 * cm, arabic(label))
        c.drawString(x0 + 0.3 * cm, y - 0.62 * cm, arabic(value))
        y -= box_h

        if y < 4 * cm:
            c.showPage()
            y = height - 2 * cm
            c.setFont(font_name, 11)

    y -= 0.8 * cm
    c.setFont(font_name, 12)
    c.drawRightString(col2, y, arabic("سعادة /"))
    y -= 0.8 * cm

    body = (
        "السلام عليكم ورحمة الله وبركاته وبعد ....\n"
        "بناءً على التنسيق المسبق فيما بين الكلية وإدارتكم الموقرة حول تدريب عدد من متدربي الكلية ضمن برنامج التدريب التعاوني،\n"
        "فإننا نوجه إليكم المتدرب الموضح بياناته أعلاه لقضاء فترة التدريب.\n"
    )

    text_obj = c.beginText()
    text_obj.setTextOrigin(x0, y)
    text_obj.setFont(font_name, 12)

    for line in body.split("\n"):
        text_obj.textLine(arabic(line))

    c.drawText(text_obj)

    c.showPage()
    c.save()
    return buf.getvalue()


# ---------------------------
# Flask App
# ---------------------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me")


LOGIN_PAGE = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8"/>
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%; background:#fff}
    .top-image img{
      width:100%;
      height:auto;
      display:block;
      object-fit:contain;
    }
    .container{
      max-width:900px;
      margin:30px auto;
      background:#fff;
      border-radius:18px;
      padding:30px;
      box-shadow:0 10px 30px rgba(0,0,0,.08);
    }
    h1{margin:10px 0 0; text-align:center; font-size:44px}
    .sub{ text-align:center; color:#444; margin-top:10px}
    form{margin-top:25px}
    .row{display:flex; gap:20px; justify-content:space-between; flex-wrap:wrap}
    .field{flex:1; min-width:240px}
    label{display:block; font-weight:700; margin:10px 0}
    input{
      width:100%;
      padding:14px 16px;
      font-size:20px;
      border:1px solid #ddd;
      border-radius:14px;
      outline:none;
    }
    button{
      width:100%;
      margin-top:20px;
      padding:18px;
      border-radius:18px;
      border:none;
      background:#0b1736;
      color:#fff;
      font-size:26px;
      cursor:pointer;
    }
    .err{color:#c40000; margin-top:14px; font-weight:700; text-align:center; white-space:pre-wrap}
    .note{color:#666; margin-top:10px; text-align:center}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>
  <div class="container">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <div class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</div>

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
      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}
      <div class="note">ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b></div>
    </form>
  </div>
</body>
</html>
"""


SELECT_PAGE = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8"/>
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%; background:#fff}
    .top-image img{width:100%; height:auto; display:block; object-fit:contain}
    .container{
      max-width:900px;
      margin:30px auto;
      background:#fff;
      border-radius:18px;
      padding:30px;
      box-shadow:0 10px 30px rgba(0,0,0,.08);
    }
    h1{margin:10px 0 0; text-align:center; font-size:40px}
    .sub{ text-align:center; color:#444; margin-top:10px}
    .chips{display:flex; gap:12px; justify-content:center; flex-wrap:wrap; margin-top:20px}
    .chip{background:#eef3ff; padding:10px 16px; border-radius:999px; font-weight:700}
    .hr{height:1px; background:#eee; margin:18px 0}
    label{display:block; font-weight:800; margin:10px 0; font-size:20px}
    select{
      width:100%;
      padding:14px 16px;
      font-size:18px;
      border:1px solid #ddd;
      border-radius:14px;
      outline:none;
      background:#fff;
    }
    button{
      width:100%;
      margin-top:18px;
      padding:18px;
      border-radius:18px;
      border:none;
      background:#0b1736;
      color:#fff;
      font-size:24px;
      cursor:pointer;
    }
    .err{color:#c40000; margin-top:14px; font-weight:800; text-align:center; white-space:pre-wrap}
    .ok{color:#0b7a2a; margin-top:14px; font-weight:800; text-align:center}
    ul{margin:0; padding:0 22px}
    li{margin:6px 0}
    a{color:#0b1736; font-weight:800}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="container">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <div class="sub">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</div>

    <div class="chips">
      <div class="chip">المتدرب: {{student["إسم المتدرب"]}}</div>
      <div class="chip">رقم المتدرب: {{student["رقم المتدرب"]}}</div>
      <div class="chip">التخصص/البرنامج: {{student["التخصص"]}} — {{student["برنامج"]}}</div>
    </div>

    <div class="hr"></div>

    {% if already %}
      <div class="ok">تم تسجيل اختيارك مسبقًا: <b>{{already}}</b></div>
      <div style="text-align:center; margin-top:12px">
        <a href="/letter?tid={{student['رقم المتدرب']}}">تحميل/طباعة خطاب التوجيه PDF</a>
      </div>
      <div class="hr"></div>
    {% endif %}

    <form method="POST" action="/choose">
      <input type="hidden" name="tid" value="{{student['رقم المتدرب']}}">
      <label>جهة التدريب المتاحة (حسب تخصصك)</label>

      <select name="org" {% if not options %}disabled{% endif %}>
        {% if options %}
          <option value="">اختر الجهة...</option>
          {% for org, rem, cap in options %}
            <option value="{{org}}">{{org}} — المتبقي {{rem}} من {{cap}}</option>
          {% endfor %}
        {% else %}
          <option>لا توجد جهات متاحة حاليًا لهذا التخصص</option>
        {% endif %}
      </select>

      <button type="submit" {% if not options %}disabled{% endif %}>حفظ الاختيار</button>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}
    </form>

    <div class="hr"></div>
    <div style="font-weight:900; font-size:20px">ملخص الفرص المتبقية داخل تخصصك:</div>
    <ul>
      {% for org, rem, cap in options %}
        <li>{{org}} : <b>{{rem}}</b> فرصة متبقية (الإجمالي {{cap}})</li>
      {% endfor %}
    </ul>
  </div>
</body>
</html>
"""


@app.get("/")
def home_get():
    return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=None)


@app.post("/")
def home_post():
    ensure_data_files()
    tid = safe_str(request.form.get("tid", ""))
    last4 = safe_str(request.form.get("last4", ""))

    try:
        df = load_students_df()
        validate_columns(df)
    except Exception as e:
        return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=f"حدث خطأ أثناء التحميل/التحقق:\n{e}")

    if not tid or not last4:
        return render_template_string(LOGIN_PAGE, title=APP_TITLE, error="يرجى إدخال الرقم التدريبي وآخر 4 أرقام من الجوال.")

    # Find student
    match = df[df["رقم المتدرب"] == tid]
    if match.empty:
        return render_template_string(LOGIN_PAGE, title=APP_TITLE, error="الرقم التدريبي غير موجود.")

    # Validate last4
    row = match.iloc[0].to_dict()
    if last4_of_phone(row.get("رقم الجوال", "")) != last4:
        return render_template_string(LOGIN_PAGE, title=APP_TITLE, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

    # Success -> redirect to select
    return redirect(url_for("select_page", tid=tid))


@app.get("/select")
def select_page():
    ensure_data_files()
    tid = safe_str(request.args.get("tid", ""))

    try:
        df = load_students_df()
        validate_columns(df)
    except Exception as e:
        return f"خطأ: {e}", 500

    match = df[df["رقم المتدرب"] == tid]
    if match.empty:
        return redirect(url_for("home_get"))

    student = match.iloc[0].to_dict()

    assignments = load_assignments()
    already = None
    if tid in assignments:
        already = assignments[tid].get("جهة التدريب", None)

    options, _summary = remaining_for_student(df, student)

    return render_template_string(
        SELECT_PAGE,
        title=APP_TITLE,
        student=student,
        options=options,
        already=already,
        error=None,
    )


@app.post("/choose")
def choose():
    ensure_data_files()
    tid = safe_str(request.form.get("tid", ""))
    org = safe_str(request.form.get("org", ""))

    try:
        df = load_students_df()
        validate_columns(df)
    except Exception as e:
        return f"خطأ: {e}", 500

    match = df[df["رقم المتدرب"] == tid]
    if match.empty:
        return redirect(url_for("home_get"))
    student = match.iloc[0].to_dict()

    assignments = load_assignments()
    if tid in assignments:
        # already chosen
        return redirect(url_for("select_page", tid=tid))

    options, _summary = remaining_for_student(df, student)
    allowed_orgs = {o for (o, rem, cap) in options}

    if not org or org not in allowed_orgs:
        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            student=student,
            options=options,
            already=None,
            error="الجهة المختارة غير متاحة الآن لهذا التخصص (قد تكون انتهت الفرص).",
        )

    # Save assignment
    assignments[tid] = {
        "رقم المتدرب": tid,
        "إسم المتدرب": student.get("إسم المتدرب", ""),
        "التخصص": student.get("التخصص", ""),
        "برنامج": student.get("برنامج", ""),
        "جهة التدريب": org,
        "timestamp": datetime.utcnow().isoformat() + "Z",
    }
    save_assignments(assignments)

    return redirect(url_for("select_page", tid=tid))


@app.get("/letter")
def letter_pdf():
    """
    Returns PDF directly (no Word, no soffice).
    """
    ensure_data_files()
    tid = safe_str(request.args.get("tid", ""))

    df = load_students_df()
    validate_columns(df)

    match = df[df["رقم المتدرب"] == tid]
    if match.empty:
        return "رقم متدرب غير صحيح.", 404

    student = match.iloc[0].to_dict()

    assignments = load_assignments()
    info = assignments.get(tid)
    if not info:
        return "لم يتم اختيار جهة تدريب بعد.", 400

    chosen_org = safe_str(info.get("جهة التدريب", ""))

    pdf_bytes = make_letter_pdf(student, chosen_org)
    return send_file(
        BytesIO(pdf_bytes),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"letter_{tid}.pdf",
    )


# ---------------------------
# Run Local / Render
# ---------------------------
if __name__ == "__main__":
    ensure_data_files()
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

import os
import json
from pathlib import Path
from collections import defaultdict, Counter

import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, render_template_string, abort
from weasyprint import HTML

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

STUDENTS_XLSX = DATA_DIR / "students.xlsx"
ASSIGNMENTS_JSON = DATA_DIR / "assignments.json"   # trainee_id -> chosen_entity
SLOTS_JSON = DATA_DIR / "slots.json"               # specialization -> entity -> remaining_count

# رابط صورة الهيدر داخل مجلد static (Flask يخدمها تلقائياً)
HEADER_IMG_URL = "/static/header.jpg"


# ---------------------------
# Helpers: JSON persistence
# ---------------------------
def read_json(path: Path, default):
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default


def write_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


# ---------------------------
# Helpers: Excel loading + column normalization
# ---------------------------
def load_students_df() -> pd.DataFrame:
    if not STUDENTS_XLSX.exists():
        raise FileNotFoundError(f"الملف غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    # نظّف أسماء الأعمدة (المشكلة عندك: "جهة التدريب " فيها مسافة)
    df.columns = [str(c).strip() for c in df.columns]

    # نظّف النصوص
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].astype(str).str.strip()

    return df


# أسماء الأعمدة حسب ملفك (كما ظهر من students.xlsx)
COL_TRAINEE_ID = "رقم المتدرب"
COL_TRAINEE_NAME = "إسم المتدرب"   # مهم: في ملفك مكتوبة "إسم" وليس "اسم"
COL_PHONE = "رقم الجوال"
COL_SPECIALIZATION = "التخصص"
COL_PROGRAM = "برنامج"
COL_TRAINING_ENTITY = "جهة التدريب"  # بعد strip تصير بدون مسافة
COL_TRAINER = "المدرب"
COL_REFNO = "الرقم المرجعي"
COL_COURSE = "اسم المقرر"
COL_DEPT = "القسم"


def ensure_required_columns(df: pd.DataFrame):
    required = [
        COL_TRAINEE_ID,
        COL_TRAINEE_NAME,
        COL_PHONE,
        COL_SPECIALIZATION,
        COL_PROGRAM,
        COL_TRAINING_ENTITY,
        COL_TRAINER,
        COL_REFNO,
        COL_COURSE,
        COL_DEPT,
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"لم أجد الأعمدة المطلوبة: {missing}\n"
            f"الأعمدة الموجودة: {list(df.columns)}"
        )


# ---------------------------
# Slots initialization from Excel
# كل جهة مكررة داخل نفس التخصص = عدد فرصها
# ---------------------------
def build_slots_from_excel(df: pd.DataFrame):
    """
    returns: dict {specialization: {entity: count}}
    """
    slots = defaultdict(lambda: defaultdict(int))
    for _, r in df.iterrows():
        spec = str(r.get(COL_SPECIALIZATION, "")).strip()
        entity = str(r.get(COL_TRAINING_ENTITY, "")).strip()
        if spec and entity and entity.lower() != "nan":
            slots[spec][entity] += 1
    # تحويل إلى dict عادي
    return {spec: dict(ent_counts) for spec, ent_counts in slots.items()}


def get_slots(df: pd.DataFrame):
    # إذا slots.json غير موجود، أنشئه من الإكسل
    slots = read_json(SLOTS_JSON, default=None)
    if not slots:
        slots = build_slots_from_excel(df)
        write_json(SLOTS_JSON, slots)
    return slots


def get_assignments():
    return read_json(ASSIGNMENTS_JSON, default={})


# ---------------------------
# UI Templates (inline)
# ---------------------------
LOGIN_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8"/>
  <title>بوابة خطاب التوجيه - التدريب التعاوني</title>
  <style>
    body{font-family: Arial, sans-serif; background:#f6f7fb; margin:0;}
    .top-image img{width:100%; max-height:240px; object-fit:cover; display:block;}
    .wrap{max-width:980px; margin:24px auto; padding:0 16px;}
    .card{background:#fff; border-radius:18px; padding:28px; box-shadow:0 8px 30px rgba(0,0,0,.08);}
    h1{margin:0 0 10px; font-size:44px; text-align:center;}
    p.sub{margin:0 0 18px; text-align:center; color:#444;}
    .grid{display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:16px;}
    label{font-weight:700;}
    input{width:100%; padding:14px 14px; border-radius:14px; border:1px solid #ddd; font-size:18px;}
    button{width:100%; margin-top:18px; padding:18px; border:0; border-radius:18px; background:#0b1220; color:#fff; font-size:22px; cursor:pointer;}
    .err{margin-top:14px; color:#c00; font-weight:700; text-align:center;}
    .note{margin-top:10px; color:#666; text-align:center;}
  </style>
</head>
<body>
  <div class="top-image"><img src="{{ header_url }}" alt="header"></div>
  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

      <form method="POST" action="/">
        <div class="grid">
          <div>
            <label>الرقم التدريبي</label>
            <input name="tid" placeholder="مثال: 444242291" required>
          </div>
          <div>
            <label>آخر 4 أرقام من الجوال</label>
            <input name="last4" placeholder="مثال: 5513" required>
          </div>
        </div>
        <button type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="err">{{ error }}</div>
      {% endif %}
      {% if note %}
        <div class="note">{{ note }}</div>
      {% endif %}
    </div>
  </div>
</body>
</html>
"""


DASHBOARD_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8"/>
  <title>اختيار جهة التدريب</title>
  <style>
    body{font-family: Arial, sans-serif; background:#f6f7fb; margin:0;}
    .top-image img{width:100%; max-height:240px; object-fit:cover; display:block;}
    .wrap{max-width:980px; margin:24px auto; padding:0 16px;}
    .card{background:#fff; border-radius:18px; padding:28px; box-shadow:0 8px 30px rgba(0,0,0,.08);}
    h1{margin:0 0 8px; font-size:44px; text-align:center;}
    p.sub{margin:0 0 18px; text-align:center; color:#444;}
    .badges{display:flex; gap:10px; flex-wrap:wrap; justify-content:center; margin:14px 0 20px;}
    .badge{background:#eef3ff; padding:10px 14px; border-radius:999px; font-weight:700;}
    select{width:100%; padding:14px; border-radius:14px; border:1px solid #ddd; font-size:18px;}
    button{width:100%; margin-top:16px; padding:18px; border:0; border-radius:18px; background:#0b1220; color:#fff; font-size:22px; cursor:pointer;}
    .err{margin-top:14px; color:#c00; font-weight:700; text-align:center;}
    .ok{margin-top:14px; color:#0a7a2a; font-weight:700; text-align:center;}
    .hr{height:1px; background:#eee; margin:22px 0;}
    ul{line-height:2; font-size:18px;}
    a{color:#0b49ff; font-weight:800;}
  </style>
</head>
<body>
  <div class="top-image"><img src="{{ header_url }}" alt="header"></div>
  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p class="sub">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

      <div class="badges">
        <div class="badge">المتدرب: {{ name }}</div>
        <div class="badge">رقم المتدرب: {{ tid }}</div>
        <div class="badge">التخصص/البرنامج: {{ spec }} — {{ program }}</div>
      </div>

      {% if already_entity %}
        <div class="ok">تم تسجيل اختيارك مسبقاً: <b>{{ already_entity }}</b></div>
        <div style="text-align:center; margin-top:10px;">
          <a href="/letter?tid={{ tid }}">تحميل/طباعة خطاب التوجيه PDF</a>
        </div>
      {% else %}
        <form method="POST" action="/choose">
          <input type="hidden" name="tid" value="{{ tid }}">
          <label style="font-weight:800;">جهة التدريب المتاحة</label>
          <select name="entity" required>
            <option value="">اختر الجهة...</option>
            {% for e, c in options %}
              <option value="{{ e }}">{{ e }} — (متبقي: {{ c }})</option>
            {% endfor %}
          </select>
          <button type="submit">حفظ الاختيار</button>
        </form>
      {% endif %}

      {% if error %}
        <div class="err">{{ error }}</div>
      {% endif %}

      <div class="hr"></div>
      <div style="font-weight:900; font-size:20px;">ملخص الفرص المتبقية داخل تخصصك:</div>
      <ul>
        {% for e, c in summary %}
          <li>{{ e }} : {{ c }} فرصة</li>
        {% endfor %}
      </ul>

    </div>
  </div>
</body>
</html>
"""


# ---------------------------
# Routes
# ---------------------------
@app.route("/", methods=["GET", "POST"])
def login():
    try:
        df = load_students_df()
        ensure_required_columns(df)
    except Exception as e:
        return render_template_string(LOGIN_HTML, header_url=HEADER_IMG_URL, error=str(e), note=None)

    if request.method == "GET":
        return render_template_string(LOGIN_HTML, header_url=HEADER_IMG_URL, error=None, note=f"ملاحظة: يتم قراءة ملف الطلاب من {STUDENTS_XLSX.as_posix()}")

    tid = str(request.form.get("tid", "")).strip()
    last4 = str(request.form.get("last4", "")).strip()

    if not tid.isdigit() or not last4.isdigit() or len(last4) != 4:
        return render_template_string(LOGIN_HTML, header_url=HEADER_IMG_URL, error="الرجاء إدخال رقم متدرب صحيح وآخر 4 أرقام (4 خانات).", note=None)

    row = df[df[COL_TRAINEE_ID].astype(str) == tid]
    if row.empty:
        return render_template_string(LOGIN_HTML, header_url=HEADER_IMG_URL, error="رقم المتدرب غير موجود.", note=None)

    phone = str(row.iloc[0][COL_PHONE]).strip()
    phone_last4 = "".join([ch for ch in phone if ch.isdigit()])[-4:]

    if phone_last4 != last4:
        return render_template_string(LOGIN_HTML, header_url=HEADER_IMG_URL, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.", note=None)

    return redirect(url_for("dashboard", tid=tid))


@app.route("/dashboard")
def dashboard():
    tid = str(request.args.get("tid", "")).strip()
    if not tid:
        return redirect(url_for("login"))

    df = load_students_df()
    ensure_required_columns(df)

    row = df[df[COL_TRAINEE_ID].astype(str) == tid]
    if row.empty:
        return redirect(url_for("login"))

    r = row.iloc[0]
    name = str(r[COL_TRAINEE_NAME]).strip()
    spec = str(r[COL_SPECIALIZATION]).strip()
    program = str(r[COL_PROGRAM]).strip()

    slots = get_slots(df)
    assignments = get_assignments()

    already_entity = assignments.get(tid)

    # خيارات تخصصه فقط + فقط المتبقي > 0
    options_dict = slots.get(spec, {})
    options = [(e, int(c)) for e, c in options_dict.items() if int(c) > 0]
    options.sort(key=lambda x: (-x[1], x[0]))

    summary = [(e, int(c)) for e, c in options_dict.items()]
    summary.sort(key=lambda x: (-x[1], x[0]))

    return render_template_string(
        DASHBOARD_HTML,
        header_url=HEADER_IMG_URL,
        tid=tid,
        name=name,
        spec=spec,
        program=program,
        options=options,
        summary=summary,
        already_entity=already_entity,
        error=None
    )


@app.route("/choose", methods=["POST"])
def choose():
    tid = str(request.form.get("tid", "")).strip()
    entity = str(request.form.get("entity", "")).strip()
    if not tid or not entity:
        return redirect(url_for("dashboard", tid=tid))

    df = load_students_df()
    ensure_required_columns(df)

    row = df[df[COL_TRAINEE_ID].astype(str) == tid]
    if row.empty:
        return redirect(url_for("login"))

    spec = str(row.iloc[0][COL_SPECIALIZATION]).strip()

    slots = get_slots(df)
    assignments = get_assignments()

    # منع التغيير إذا سبق اختار (حسب طلبك لتقليل التلاعب)
    if tid in assignments:
        return redirect(url_for("dashboard", tid=tid))

    # تحقق أن الجهة ضمن تخصصه وأنها متاحة
    spec_slots = slots.get(spec, {})
    remaining = int(spec_slots.get(entity, 0))
    if remaining <= 0:
        return render_template_string(
            DASHBOARD_HTML,
            header_url=HEADER_IMG_URL,
            tid=tid,
            name=str(row.iloc[0][COL_TRAINEE_NAME]).strip(),
            spec=spec,
            program=str(row.iloc[0][COL_PROGRAM]).strip(),
            options=[(e, int(c)) for e, c in spec_slots.items() if int(c) > 0],
            summary=sorted([(e, int(c)) for e, c in spec_slots.items()], key=lambda x: (-x[1], x[0])),
            already_entity=None,
            error="هذه الجهة غير متاحة حالياً (انتهت الفرص)."
        )

    # احفظ الاختيار + أنقص الفرصة
    assignments[tid] = entity
    spec_slots[entity] = remaining - 1
    slots[spec] = spec_slots

    write_json(ASSIGNMENTS_JSON, assignments)
    write_json(SLOTS_JSON, slots)

    return redirect(url_for("dashboard", tid=tid))


# ---------------------------
# PDF Generation (NO soffice)
# ---------------------------
def build_letter_html(student: dict, chosen_entity: str) -> str:
    # ملاحظة: هذا قالب مرتب قريب من نموذجك (جدول + بيانات)
    # إذا تبغى نفس شكل قالبك 100%، أرسل لقطة واضحة للصفحة الأولى من "قالب الورد" أو PDF القالب وسأطابقه بالـ CSS.
    return f"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <style>
    @page {{
      size: A4;
      margin: 18mm 16mm 18mm 16mm;
    }}
    body {{
      font-family: "DejaVu Sans", Arial, sans-serif;
      font-size: 14px;
      color: #000;
    }}
    .header img {{
      width: 100%;
      height: auto;
      max-height: 120px;
      object-fit: contain;
      display: block;
      margin-bottom: 10px;
    }}
    h2 {{
      text-align: center;
      margin: 10px 0 14px;
      font-size: 22px;
    }}
    table {{
      width: 100%;
      border-collapse: collapse;
      margin: 10px 0 14px;
      table-layout: fixed;
    }}
    th, td {{
      border: 1px solid #000;
      padding: 8px 10px;
      vertical-align: middle;
      word-wrap: break-word;
    }}
    th {{
      background: #e9e9e9;
      width: 28%;
      text-align: right;
      font-weight: 700;
    }}
    td {{
      text-align: right;
    }}
    .content {{
      margin-top: 6px;
      line-height: 1.9;
      text-align: justify;
    }}
  </style>
</head>
<body>
  <div class="header">
    <img src="{HEADER_IMG_URL}">
  </div>

  <h2>خطاب توجيه متدرب تدريب تعاوني</h2>

  <table>
    <tr><th>الرقم</th><td>{student.get("رقم المتدرب","")}</td></tr>
    <tr><th>الاسم</th><td>{student.get("إسم المتدرب","")}</td></tr>
    <tr><th>الرقم الأكاديمي</th><td>{student.get("رقم المتدرب","")}</td></tr>
    <tr><th>التخصص</th><td>{student.get("التخصص","")}</td></tr>
    <tr><th>جوال</th><td>{student.get("رقم الجوال","")}</td></tr>
    <tr><th>اسم المشرف من الكلية</th><td>{student.get("المدرب","")}</td></tr>
    <tr><th>الرقم المرجعي للمقرر</th><td>{student.get("الرقم المرجعي","")}</td></tr>
    <tr><th>جهة التدريب المختارة</th><td>{chosen_entity}</td></tr>
  </table>

  <div class="content">
    <div><b>سعادة /</b></div>
    <div>السلام عليكم ورحمة الله وبركاته وبعد ....</div>
    <div>
      بناءً على التنسيق المسبق فيما بين الكلية وإدارتكم الموقرة حول تدريب عدد من متدربي الكلية ضمن برنامج التدريب التعاوني،
      فإننا نوجه إليكم المتدرب الموضح بياناته أعلاه لقضاء فترة التدريب.
    </div>
    <div style="margin-top:8px;">
      وتفضلوا بقبول فائق الاحترام والتقدير.
    </div>
  </div>
</body>
</html>
"""


@app.route("/letter")
def letter_pdf():
    tid = str(request.args.get("tid", "")).strip()
    if not tid:
        abort(404)

    df = load_students_df()
    ensure_required_columns(df)

    row = df[df[COL_TRAINEE_ID].astype(str) == tid]
    if row.empty:
        abort(404)

    assignments = get_assignments()
    chosen_entity = assignments.get(tid)
    if not chosen_entity:
        return "لم يتم تسجيل جهة تدريب لهذا المتدرب بعد.", 400

    # بيانات الطالب من الإكسل
    r = row.iloc[0].to_dict()
    # تأكد من وجود keys الأساسية كما في الإكسل
    student = {
        "رقم المتدرب": str(r.get(COL_TRAINEE_ID, "")).strip(),
        "إسم المتدرب": str(r.get(COL_TRAINEE_NAME, "")).strip(),
        "رقم الجوال": str(r.get(COL_PHONE, "")).strip(),
        "التخصص": str(r.get(COL_SPECIALIZATION, "")).strip(),
        "برنامج": str(r.get(COL_PROGRAM, "")).strip(),
        "المدرب": str(r.get(COL_TRAINER, "")).strip(),
        "الرقم المرجعي": str(r.get(COL_REFNO, "")).strip(),
    }

    html = build_letter_html(student, chosen_entity)

    # WeasyPrint يحتاج base_url عشان يقرأ /static/header.jpg
    pdf_bytes = HTML(string=html, base_url=str(BASE_DIR)).write_pdf()

    out_path = DATA_DIR / f"letter_{tid}.pdf"
    out_path.write_bytes(pdf_bytes)

    return send_file(
        out_path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"خطاب_توجيه_{tid}.pdf",
    )


# ---------------------------
# Run local
# ---------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

import os
import io
import re
from datetime import datetime

import pandas as pd
from flask import Flask, request, redirect, url_for, render_template_string, send_file

from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


# =========================================================
# Flask App (مهم جدًا يكون هنا قبل أي @app.route)
# =========================================================
app = Flask(__name__)

# =========================================================
# Paths
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data", "students.xlsx")
ASSIGNMENTS_FILE = os.path.join(BASE_DIR, "data", "assignments.json")

# =========================================================
# Helpers
# =========================================================
REQUIRED_COLS = [
    "رقم المتدرب",
    "اسم المتدرب",
    "برنامج",
    "رقم الجوال",
    "جهة التدريب",
    "المدرب",
    "الرقم المرجعي",
    "اسم المقرر",
]


def normalize_digits(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip()
    # تحويل الأرقام العربية إلى إنجليزية
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")
    s = s.translate(trans)
    # إزالة أي شيء غير رقم
    s = re.sub(r"\D+", "", s)
    return s


def load_students() -> pd.DataFrame:
    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(
            f"لم أجد الأعمدة المطلوبة: {missing}. الأعمدة الموجودة: {list(df.columns)}"
        )

    # توحيد الأنواع
    df["رقم المتدرب"] = df["رقم المتدرب"].apply(normalize_digits)
    df["رقم الجوال"] = df["رقم الجوال"].apply(normalize_digits)

    # برنامج/جهة/اسم
    for col in ["اسم المتدرب", "برنامج", "جهة التدريب", "المدرب", "الرقم المرجعي", "اسم المقرر"]:
        df[col] = df[col].astype(str).fillna("").apply(lambda x: x.strip())

    return df


def last4(phone: str) -> str:
    p = normalize_digits(phone)
    return p[-4:] if len(p) >= 4 else ""


# =========================================================
# HTML Templates (بسيطة)
# =========================================================
LOGIN_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>بوابة خطاب التوجيه - التدريب التعاوني</title>
<style>
body{font-family:Tahoma,Arial; background:#f6f7fb; margin:0}
.header img{width:100%; height:auto; display:block}
.card{max-width:900px; margin:30px auto; background:#fff; border-radius:18px; padding:28px; box-shadow:0 10px 30px rgba(0,0,0,.08)}
h1{margin:0 0 10px; font-size:42px; text-align:center}
p{margin:0 0 18px; text-align:center; color:#444}
.form-row{display:flex; gap:18px; margin-top:22px}
.form-row>div{flex:1}
label{display:block; font-weight:700; margin-bottom:10px}
input{width:100%; padding:14px; border:1px solid #ddd; border-radius:14px; font-size:18px}
button{width:100%; margin-top:18px; padding:16px; border:0; border-radius:18px; background:#0b1630; color:#fff; font-size:20px; cursor:pointer}
.err{margin-top:14px; color:#b30000; font-weight:700; text-align:center}
.note{margin-top:10px; color:#666; text-align:center}
</style>
</head>
<body>
<div class="header">
  <img src="/static/header.jpg" alt="Header">
</div>
<div class="card">
  <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
  <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>
  <form method="post" action="/">
    <div class="form-row">
      <div>
        <label>الرقم التدريبي</label>
        <input name="tid" placeholder="مثال: 444229747" required>
      </div>
      <div>
        <label>آخر 4 أرقام من الجوال</label>
        <input name="last4" placeholder="مثال: 6101" required>
      </div>
    </div>
    <button type="submit">دخول</button>
  </form>

  {% if error %}<div class="err">{{error}}</div>{% endif %}
  <div class="note">ملاحظة: يتم قراءة ملف الطلاب من data/students.xlsx</div>
</div>
</body>
</html>
"""


PROFILE_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>اختيار جهة التدريب</title>
<style>
body{font-family:Tahoma,Arial; background:#f6f7fb; margin:0}
.header img{width:100%; height:auto; display:block}
.card{max-width:1000px; margin:30px auto; background:#fff; border-radius:18px; padding:28px; box-shadow:0 10px 30px rgba(0,0,0,.08)}
h1{margin:0 0 10px; font-size:34px}
.badges{display:flex; gap:12px; flex-wrap:wrap; margin:18px 0}
.badge{background:#eef3ff; padding:10px 14px; border-radius:999px; font-weight:700}
select{width:100%; padding:14px; border:1px solid #ddd; border-radius:14px; font-size:18px}
button{width:100%; margin-top:18px; padding:16px; border:0; border-radius:18px; background:#0b1630; color:#fff; font-size:20px; cursor:pointer}
.err{margin-top:14px; color:#b30000; font-weight:700; text-align:center}
.ok{margin-top:14px; color:#0a7a2d; font-weight:700; text-align:center}
a{font-weight:700}
</style>
</head>
<body>
<div class="header">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="card">
  <h1>اختيار جهة التدريب</h1>

  <div class="badges">
    <div class="badge">المتدرب: {{name}}</div>
    <div class="badge">رقم المتدرب: {{tid}}</div>
    <div class="badge">البرنامج: {{program}}</div>
  </div>

  {% if chosen %}
    <div class="ok">تم تسجيل اختيارك مسبقًا: <b>{{chosen}}</b></div>
    <p style="text-align:center; margin-top:10px;">
      <a href="/letter?tid={{tid}}">تحميل خطاب التوجيه PDF</a>
    </p>
  {% else %}
    <form method="post" action="/choose">
      <input type="hidden" name="tid" value="{{tid}}">
      <label style="font-weight:700; display:block; margin-bottom:10px;">جهة التدريب المتاحة</label>
      <select name="org" required>
        <option value="">اختر الجهة...</option>
        {% for o in options %}
          <option value="{{o}}">{{o}}</option>
        {% endfor %}
      </select>
      <button type="submit">حفظ الاختيار</button>
    </form>
    {% if error %}<div class="err">{{error}}</div>{% endif %}
    {% if options|length == 0 %}
      <div class="err">لا توجد جهات متاحة حاليًا لهذا البرنامج/التخصص.</div>
    {% endif %}
  {% endif %}
</div>
</body>
</html>
"""


# =========================================================
# Routes
# =========================================================
@app.route("/", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        tid = normalize_digits(request.form.get("tid", ""))
        l4 = normalize_digits(request.form.get("last4", ""))

        try:
            df = load_students()
        except Exception as e:
            return render_template_string(LOGIN_HTML, error=str(e))

        row = df[df["رقم المتدرب"] == tid]
        if row.empty:
            error = "الرقم التدريبي غير صحيح."
        else:
            phone = row.iloc[0]["رقم الجوال"]
            if last4(phone) != l4:
                error = "آخر 4 أرقام من الجوال غير صحيحة."
            else:
                return redirect(url_for("choose_page", tid=tid))

    return render_template_string(LOGIN_HTML, error=error)


@app.route("/choose", methods=["GET"])
def choose_page():
    tid = normalize_digits(request.args.get("tid", ""))

    df = load_students()
    row = df[df["رقم المتدرب"] == tid]
    if row.empty:
        return redirect(url_for("login"))

    row0 = row.iloc[0]
    name = row0["اسم المتدرب"]
    program = row0["برنامج"]

    # الجهات المتاحة من ملف excel لنفس البرنامج فقط
    df_same = df[df["برنامج"] == program].copy()

    # كل جهة عدد فرصها = عدد تكرارها (كما طلبت)
    counts = df_same["جهة التدريب"].value_counts().to_dict()

    # قراءة الاختيارات السابقة (نحفظها في ملف JSON داخل data)
    chosen_map = {}
    if os.path.exists(ASSIGNMENTS_FILE):
        import json
        with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as f:
            chosen_map = json.load(f)

    chosen = chosen_map.get(tid)

    # إن لم يختَر سابقًا، نعرض فقط الجهات التي ما زال لها فرص (عدد الاختيارات < العدد)
    taken_counts = {}
    for _tid, _org in chosen_map.items():
        taken_counts[_org] = taken_counts.get(_org, 0) + 1

    options = []
    for org, total in counts.items():
        taken = taken_counts.get(org, 0)
        if taken < total:
            options.append(org)

    # ترتيب
    options = sorted(options)

    return render_template_string(
        PROFILE_HTML,
        tid=tid,
        name=name,
        program=program,
        options=options,
        chosen=chosen,
        error=None,
    )


@app.route("/choose", methods=["POST"])
def save_choice():
    tid = normalize_digits(request.form.get("tid", ""))
    org = (request.form.get("org", "") or "").strip()

    df = load_students()
    row = df[df["رقم المتدرب"] == tid]
    if row.empty:
        return redirect(url_for("login"))

    program = row.iloc[0]["برنامج"]
    df_same = df[df["برنامج"] == program].copy()
    counts = df_same["جهة التدريب"].value_counts().to_dict()

    if org not in counts:
        return render_template_string(PROFILE_HTML,
                                      tid=tid,
                                      name=row.iloc[0]["اسم المتدرب"],
                                      program=program,
                                      options=[],
                                      chosen=None,
                                      error="الجهة المختارة غير صالحة لهذا البرنامج.")

    # تحميل الاختيارات
    chosen_map = {}
    if os.path.exists(ASSIGNMENTS_FILE):
        import json
        with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as f:
            chosen_map = json.load(f)

    # إذا سبق اختار
    if tid in chosen_map:
        return redirect(url_for("choose_page", tid=tid))

    # تحقق من السعة
    taken_counts = {}
    for _tid, _org in chosen_map.items():
        taken_counts[_org] = taken_counts.get(_org, 0) + 1

    if taken_counts.get(org, 0) >= counts[org]:
        return render_template_string(PROFILE_HTML,
                                      tid=tid,
                                      name=row.iloc[0]["اسم المتدرب"],
                                      program=program,
                                      options=[],
                                      chosen=None,
                                      error="عذرًا، اكتملت فرص هذه الجهة. اختر جهة أخرى.")

    # حفظ
    chosen_map[tid] = org
    import json
    os.makedirs(os.path.dirname(ASSIGNMENTS_FILE), exist_ok=True)
    with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(chosen_map, f, ensure_ascii=False, indent=2)

    return redirect(url_for("choose_page", tid=tid))


@app.route("/letter", methods=["GET"])
def letter_pdf():
    tid = normalize_digits(request.args.get("tid", ""))

    df = load_students()
    row = df[df["رقم المتدرب"] == tid]
    if row.empty:
        return "رقم المتدرب غير موجود", 404
    r = row.iloc[0]

    # قراءة جهة الاختيار
    chosen_map = {}
    if os.path.exists(ASSIGNMENTS_FILE):
        import json
        with open(ASSIGNMENTS_FILE, "r", encoding="utf-8") as f:
            chosen_map = json.load(f)
    org = chosen_map.get(tid)
    if not org:
        return "لم يتم اختيار جهة تدريب بعد.", 400

    # ========= PDF (ReportLab) =========
    buffer = io.BytesIO()

    # محاولة تسجيل خط عربي (إن توفر داخل المشروع ضع خط TTF داخل static مثلاً)
    # إن لم يوجد، سيعمل بخط افتراضي لكنه قد لا يدعم العربي بشكل ممتاز.
    font_path = os.path.join(BASE_DIR, "static", "Cairo-Regular.ttf")
    if os.path.exists(font_path):
        pdfmetrics.registerFont(TTFont("Cairo", font_path))
        font_name = "Cairo"
    else:
        font_name = "Helvetica"

    c = canvas.Canvas(buffer, pagesize=A4)
    w, h = A4

    c.setFont(font_name, 16)
    c.drawRightString(w - 50, h - 80, "خطاب توجيه متدرب تدريب تعاوني")

    c.setFont(font_name, 12)
    y = h - 130

    lines = [
        f"الرقم: {r['رقم المتدرب']}",
        f"الاسم: {r['اسم المتدرب']}",
        f"الرقم الأكاديمي: {r['رقم المتدرب']}",
        f"التخصص/البرنامج: {r['برنامج']}",
        f"جوال: {r['رقم الجوال']}",
        f"اسم المشرف من الكلية: {r['المدرب']}",
        f"الرقم المرجعي للمقرر: {r['الرقم المرجعي']}",
        f"جهة التدريب المختارة: {org}",
    ]

    for line in lines:
        c.drawRightString(w - 50, y, line)
        y -= 22

    c.setFont(font_name, 12)
    c.drawRightString(w - 50, y - 20, "السادة/ ........................................")
    c.drawRightString(w - 50, y - 45, "السلام عليكم ورحمة الله وبركاته وبعد ...")

    c.showPage()
    c.save()

    buffer.seek(0)
    filename = f"خطاب_توجيه_{tid}.pdf"
    return send_file(buffer, as_attachment=True, download_name=filename, mimetype="application/pdf")


# =========================================================
# Local run
# =========================================================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)
________________________________________
✅ الآن نفّذ هذه الخطوات بالضبط على GitHub
1.	افتح app.py → اضغط ✏️ Edit
2.	احذف كل شيء → الصق الكود كامل → Commit
3.	افتح requirements.txt → الصقه بالكامل:
flask==3.0.2
gunicorn==22.0.0
pandas==2.2.2
openpyxl==3.1.5
python-docx==1.1.2
reportlab==4.2.5
ثم Commit
4.	في Render: Settings → Build & Deploy → Start Command:
gunicorn app:app --bind 0.0.0.0:$PORT



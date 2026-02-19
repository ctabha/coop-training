import os
import re
import json
import uuid
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import (
    Flask,
    request,
    render_template_string,
    redirect,
    url_for,
    session,
    send_file,
)

# =========================
# إعدادات عامة
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

DATA_FILE = DATA_DIR / "students.xlsx"   # <-- مهم: نفس اسم الملف داخل data
SLOTS_FILE = DATA_DIR / "slots.json"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"

# مفاتيح الأعمدة المتوقعة في الإكسل (بأسماءك الحالية)
COL_TRAINEE_ID = "رقم المتدرب"
COL_TRAINEE_NAME = "إسم المتدرب"
COL_PHONE = "رقم الجوال"
COL_SPECIALTY = "التخصص"
COL_DEPT = "القسم"
COL_PROGRAM = "برنامج"
COL_TRAINING_ORG = "جهة التدريب"  # في ملفك يظهر "جهة التدريب " بمسافة.. سننظّفها

# =========================
# Flask App
# =========================
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")


# =========================
# أدوات مساعدة
# =========================
def norm_colname(s: str) -> str:
    """توحيد أسماء الأعمدة: إزالة مسافات زائدة وتبديل المسافات الداخلية."""
    if s is None:
        return ""
    s = str(s).strip()
    s = re.sub(r"\s+", " ", s)
    return s


def digits_only(x) -> str:
    """استخراج الأرقام فقط من أي قيمة."""
    if x is None:
        return ""
    s = str(x)
    return "".join(ch for ch in s if ch.isdigit())


def last4_phone(x) -> str:
    d = digits_only(x)
    return d[-4:] if len(d) >= 4 else d


def safe_read_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def safe_write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def find_students_sheet(excel_path: Path) -> str:
    """
    يختار الشيت الصحيح تلقائياً:
    أول شيت يحتوي عمود 'رقم المتدرب' بعد تنظيف أسماء الأعمدة.
    """
    xls = pd.ExcelFile(excel_path)
    for sh in xls.sheet_names:
        df0 = pd.read_excel(excel_path, sheet_name=sh, nrows=2)
        cols = [norm_colname(c) for c in df0.columns.tolist()]
        if COL_TRAINEE_ID in cols:
            return sh
    # إذا ما لقينا، نرجع أول شيت
    return xls.sheet_names[0]


def load_students() -> pd.DataFrame:
    """
    يقرأ students.xlsx ويعيد DataFrame بأعمدة نظيفة.
    """
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {DATA_FILE}")

    sheet = find_students_sheet(DATA_FILE)
    df = pd.read_excel(DATA_FILE, sheet_name=sheet)

    # تنظيف أسماء الأعمدة (المسافات الزائدة تسبب مشاكل مثل 'جهة التدريب ')
    df.columns = [norm_colname(c) for c in df.columns]

    # تأكد من وجود الأعمدة الأساسية
    required = {COL_TRAINEE_ID, COL_TRAINEE_NAME, COL_PHONE, COL_SPECIALTY, COL_TRAINING_ORG}
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"أعمدة ناقصة في ملف الإكسل: {missing}\n"
            f"الأعمدة الموجودة: {df.columns.tolist()}"
        )

    # توحيد بعض الأعمدة كسلاسل
    df[COL_TRAINEE_ID] = df[COL_TRAINEE_ID].apply(digits_only)
    df[COL_PHONE] = df[COL_PHONE].apply(digits_only)
    df[COL_SPECIALTY] = df[COL_SPECIALTY].astype(str).str.strip()
    df[COL_TRAINING_ORG] = df[COL_TRAINING_ORG].astype(str).str.strip()

    # حذف صفوف غير مفيدة
    df = df[df[COL_TRAINEE_ID] != ""]
    df = df[df[COL_SPECIALTY] != ""]
    df = df[df[COL_TRAINING_ORG] != ""]

    return df


def build_slots_from_excel(df: pd.DataFrame) -> dict:
    """
    يبني الفرص حسب التخصص والجهة من تكرار الجهة داخل الإكسل:
    كل (تخصص + جهة) عدد التكرارات = عدد الفرص المتاحة.
    """
    g = df.groupby([COL_SPECIALTY, COL_TRAINING_ORG]).size().reset_index(name="count")
    slots = {}
    for _, row in g.iterrows():
        spec = str(row[COL_SPECIALTY]).strip()
        org = str(row[COL_TRAINING_ORG]).strip()
        cnt = int(row["count"])
        slots.setdefault(spec, {})
        slots[spec][org] = cnt
    return slots


def ensure_slots(df: pd.DataFrame) -> dict:
    """
    إذا slots.json موجود نستخدمه، وإلا ننشئه من الإكسل.
    """
    slots = safe_read_json(SLOTS_FILE, default=None)
    if not isinstance(slots, dict) or not slots:
        slots = build_slots_from_excel(df)
        safe_write_json(SLOTS_FILE, slots)
    return slots


def ensure_assignments() -> dict:
    """
    assignments.json: تخزين اختيار كل متدرب:
    { "رقم_المتدرب": {"org": "...", "ts": "..."} }
    """
    a = safe_read_json(ASSIGNMENTS_FILE, default={})
    if not isinstance(a, dict):
        a = {}
    return a


def get_student_record(df: pd.DataFrame, trainee_id: str):
    rows = df[df[COL_TRAINEE_ID] == trainee_id]
    if rows.empty:
        return None
    # لو تكرر، نأخذ أول صف
    return rows.iloc[0].to_dict()


def available_orgs_for_specialty(slots: dict, specialty: str):
    """
    يرجع الجهات المتاحة للتخصص (اللي رصيدها > 0)
    """
    spec_map = slots.get(specialty, {})
    items = [(org, int(cnt)) for org, cnt in spec_map.items() if int(cnt) > 0]
    # ترتيب تنازلي حسب الفرص
    items.sort(key=lambda x: x[1], reverse=True)
    return items


# =========================
# واجهات HTML
# =========================
PAGE_LOGIN = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
  body{font-family:Arial;background:#f7f7f7;margin:0}
  .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
  .top-image img{
    width:100%;
    height:100%;
    object-fit:contain; /* يمنع القص */
    object-position:center;
    display:block;
  }
  .wrap{max-width:980px;margin:24px auto;padding:0 16px}
  .card{
    background:#fff;border-radius:18px;padding:28px;
    box-shadow:0 12px 30px rgba(0,0,0,.08);
    text-align:center;
  }
  h1{margin:0 0 10px;font-size:44px}
  p{margin:0 0 22px;color:#444}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin:10px 0 18px}
  label{display:block;text-align:right;font-weight:700;margin:0 0 6px}
  input{
    width:100%;padding:14px 16px;border:1px solid #ddd;border-radius:14px;
    font-size:18px;outline:none;
  }
  .btn{
    width:100%;padding:18px;border:0;border-radius:18px;
    background:#0b1730;color:#fff;font-size:22px;cursor:pointer;
  }
  .err{color:#c00;margin-top:14px;font-weight:700}
  .note{color:#666;margin-top:10px;font-size:14px}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="wrap">
  <div class="card">
    <h1>{{title}}</h1>
    <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

    <form method="post" action="{{ url_for('login') }}">
      <div class="grid">
        <div>
          <label>الرقم التدريبي</label>
          <input name="trainee_id" placeholder="مثال: 444229747" required>
        </div>
        <div>
          <label>آخر 4 أرقام من الجوال</label>
          <input name="last4" placeholder="مثال: 6101" required>
        </div>
      </div>
      <button class="btn" type="submit">دخول</button>
    </form>

    {% if error %}
      <div class="err">{{error}}</div>
    {% endif %}

    {% if load_error %}
      <div class="err">حدث خطأ أثناء التحميل/التحقق: {{load_error}}</div>
      <div class="note">تأكد أن الملف موجود داخل <b>data/students.xlsx</b> وأن الأعمدة مطابقة.</div>
    {% endif %}
  </div>
</div>

</body>
</html>
"""

PAGE_CHOOSE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
  body{font-family:Arial;background:#f7f7f7;margin:0}
  .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
  .top-image img{width:100%;height:100%;object-fit:contain;display:block}
  .wrap{max-width:980px;margin:24px auto;padding:0 16px}
  .card{
    background:#fff;border-radius:18px;padding:28px;
    box-shadow:0 12px 30px rgba(0,0,0,.08);
  }
  h1{margin:0 0 6px;font-size:34px;text-align:center}
  .sub{margin:0 0 18px;text-align:center;color:#444}
  .row{display:flex;gap:12px;flex-wrap:wrap;justify-content:space-between;margin-bottom:14px}
  .pill{background:#eef2ff;border-radius:999px;padding:10px 14px;font-weight:700}
  select{
    width:100%;padding:14px 16px;border:1px solid #ddd;border-radius:14px;
    font-size:16px;outline:none;background:#fff;
  }
  .btn{
    width:100%;padding:16px;border:0;border-radius:16px;
    background:#0b1730;color:#fff;font-size:18px;cursor:pointer;margin-top:12px;
  }
  .err{color:#c00;margin-top:14px;font-weight:700;text-align:center}
  .ok{color:#0a7a2f;margin-top:14px;font-weight:700;text-align:center}
  .small{color:#666;font-size:13px;text-align:center;margin-top:8px}
  .hr{height:1px;background:#eee;margin:18px 0}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="wrap">
  <div class="card">
    <h1>{{title}}</h1>
    <p class="sub">اختر جهة التدريب المتاحة لتخصصك ثم اطبع خطاب التوجيه.</p>

    <div class="row">
      <div class="pill">المتدرب: {{ trainee_name }}</div>
      <div class="pill">رقم المتدرب: {{ trainee_id }}</div>
      <div class="pill">التخصص: {{ specialty }}</div>
    </div>

    <div class="hr"></div>

    {% if already_org %}
      <p class="ok">تم اختيار جهة التدريب مسبقًا: <b>{{already_org}}</b></p>
      <form method="post" action="{{ url_for('print_letter') }}">
        <button class="btn" type="submit">طباعة خطاب التوجيه PDF</button>
      </form>
      <div class="small">إذا تبغى تغيير الجهة: احذف ملف assignments.json أو احذف تسجيل المتدرب منه.</div>

    {% else %}
      <form method="post" action="{{ url_for('choose_org') }}">
        <label style="font-weight:700;display:block;margin-bottom:6px">جهة التدريب المتاحة</label>
        <select name="org" required>
          <option value="" disabled selected>اختر الجهة...</option>
          {% for org, cnt in orgs %}
            <option value="{{org}}">{{org}} (متبقي: {{cnt}})</option>
          {% endfor %}
        </select>
        <button class="btn" type="submit">حفظ الاختيار</button>
      </form>
      {% if orgs|length == 0 %}
        <p class="err">لا توجد جهات متاحة حاليًا لهذا التخصص.</p>
      {% endif %}
    {% endif %}

    {% if error %}
      <p class="err">{{error}}</p>
    {% endif %}

  </div>
</div>

</body>
</html>
"""


# =========================
# Routes
# =========================
@app.route("/", methods=["GET"])
def home():
    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=None, load_error=None)


@app.route("/", methods=["POST"])
def login():
    trainee_id = digits_only(request.form.get("trainee_id", ""))
    l4 = digits_only(request.form.get("last4", ""))

    try:
        df = load_students()
    except Exception as e:
        return render_template_string(
            PAGE_LOGIN, title=APP_TITLE, error=None, load_error=str(e)
        )

    rec = get_student_record(df, trainee_id)
    if not rec:
        return render_template_string(
            PAGE_LOGIN, title=APP_TITLE, error="الرقم التدريبي غير موجود.", load_error=None
        )

    # تحقق آخر 4 أرقام من الجوال
    phone_last4 = last4_phone(rec.get(COL_PHONE, ""))
    if l4 != phone_last4:
        return render_template_string(
            PAGE_LOGIN,
            title=APP_TITLE,
            error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.",
            load_error=None,
        )

    # حفظ في session
    session["trainee_id"] = trainee_id
    session["trainee_name"] = str(rec.get(COL_TRAINEE_NAME, "")).strip()
    session["specialty"] = str(rec.get(COL_SPECIALTY, "")).strip()
    session["program"] = str(rec.get(COL_PROGRAM, "")).strip()
    session["dept"] = str(rec.get(COL_DEPT, "")).strip()

    return redirect(url_for("choose_org"))


@app.route("/choose", methods=["GET"])
def choose_org():
    trainee_id = session.get("trainee_id")
    if not trainee_id:
        return redirect(url_for("home"))

    try:
        df = load_students()
        slots = ensure_slots(df)
        assignments = ensure_assignments()
    except Exception as e:
        return render_template_string(
            PAGE_LOGIN, title=APP_TITLE, error=None, load_error=str(e)
        )

    specialty = session.get("specialty", "").strip()
    orgs = available_orgs_for_specialty(slots, specialty)

    already = assignments.get(trainee_id, {}).get("org")

    return render_template_string(
        PAGE_CHOOSE,
        title=APP_TITLE,
        trainee_id=trainee_id,
        trainee_name=session.get("trainee_name", ""),
        specialty=specialty,
        orgs=orgs,
        already_org=already,
        error=None,
    )


@app.route("/choose", methods=["POST"])
def choose_org_post():
    trainee_id = session.get("trainee_id")
    if not trainee_id:
        return redirect(url_for("home"))

    chosen = (request.form.get("org") or "").strip()
    if not chosen:
        return redirect(url_for("choose_org"))

    df = load_students()
    slots = ensure_slots(df)
    assignments = ensure_assignments()

    specialty = session.get("specialty", "").strip()

    # منع إعادة الاختيار لو سبق اختار
    if trainee_id in assignments:
        return redirect(url_for("choose_org"))

    # تحقق أن الجهة متاحة ورصيدها > 0
    spec_map = slots.get(specialty, {})
    if chosen not in spec_map or int(spec_map.get(chosen, 0)) <= 0:
        return render_template_string(
            PAGE_CHOOSE,
            title=APP_TITLE,
            trainee_id=trainee_id,
            trainee_name=session.get("trainee_name", ""),
            specialty=specialty,
            orgs=available_orgs_for_specialty(slots, specialty),
            already_org=None,
            error="هذه الجهة غير متاحة الآن (قد تكون اكتملت الفرص). اختر جهة أخرى.",
        )

    # خصم فرصة
    spec_map[chosen] = int(spec_map[chosen]) - 1
    slots[specialty] = spec_map
    safe_write_json(SLOTS_FILE, slots)

    # حفظ اختيار المتدرب
    assignments[trainee_id] = {
        "org": chosen,
        "ts": datetime.now().isoformat(timespec="seconds"),
        "specialty": specialty,
        "name": session.get("trainee_name", ""),
    }
    safe_write_json(ASSIGNMENTS_FILE, assignments)

    return redirect(url_for("choose_org"))


# alias route names to match url_for calls above
choose_org.endpoint = "choose_org"
choose_org_post.endpoint = "choose_org"


# =========================
# طباعة PDF (قابل للتخصيص)
# =========================
def replace_in_docx(doc, mapping: dict):
    """
    استبدال نصوص بسيطة داخل docx (في البراجرافات والجداول)
    """
    def repl_text(text: str) -> str:
        for k, v in mapping.items():
            text = text.replace(k, v)
        return text

    for p in doc.paragraphs:
        for run in p.runs:
            if run.text:
                run.text = repl_text(run.text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        if run.text:
                            run.text = repl_text(run.text)


@app.route("/print", methods=["POST"])
def print_letter():
    trainee_id = session.get("trainee_id")
    if not trainee_id:
        return redirect(url_for("home"))

    template_path = DATA_DIR / "letter_template.docx"
    if not template_path.exists():
        return f"ملف القالب غير موجود: {template_path}", 400

    assignments = ensure_assignments()
    chosen_org = assignments.get(trainee_id, {}).get("org")
    if not chosen_org:
        return redirect(url_for("choose_org"))

    # تجهيز بيانات الخطاب
    mapping = {
        "{{TRAINEE_NAME}}": session.get("trainee_name", ""),
        "{{TRAINEE_ID}}": trainee_id,
        "{{SPECIALTY}}": session.get("specialty", ""),
        "{{PROGRAM}}": session.get("program", ""),
        "{{ORG_NAME}}": chosen_org,
        "{{DATE}}": datetime.now().strftime("%Y-%m-%d"),
    }

    # إنشاء ملفات مؤقتة
    workdir = BASE_DIR / "out"
    workdir.mkdir(parents=True, exist_ok=True)
    tmp_id = uuid.uuid4().hex
    tmp_docx = workdir / f"letter_{tmp_id}.docx"

    # تعديل القالب
    from docx import Document
    doc = Document(str(template_path))
    replace_in_docx(doc, mapping)
    doc.save(str(tmp_docx))

    # تحويل إلى PDF عبر LibreOffice (في الدوكر عندك موجود)
    tmp_pdf = workdir / f"letter_{tmp_id}.pdf"
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(workdir),
        str(tmp_docx),
    ]
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if p.returncode != 0:
        return f"فشل تحويل PDF:\n{p.stderr}\n{p.stdout}", 500

    # LibreOffice يطلع اسم pdf نفس docx غالباً
    produced_pdf = workdir / f"letter_{tmp_id}.pdf"
    if not produced_pdf.exists():
        # fallback: يبحث عن أقرب ملف pdf
        pdfs = sorted(workdir.glob("letter_*.pdf"), key=lambda x: x.stat().st_mtime, reverse=True)
        if pdfs:
            produced_pdf = pdfs[0]
        else:
            return "لم يتم العثور على ملف PDF الناتج.", 500

    return send_file(
        str(produced_pdf),
        as_attachment=True,
        download_name="خطاب_التوجيه.pdf",
        mimetype="application/pdf",
    )


# اسم endpoint للزر في الصفحة
print_letter.endpoint = "print_letter"


# =========================
# تشغيل محلي
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

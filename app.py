import os
import json
from datetime import datetime
from pathlib import Path
import subprocess

import pandas as pd
from flask import (
    Flask, request, redirect, url_for,
    render_template_string, send_file, abort
)
from docx import Document

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

# ملفات البيانات
STUDENTS_CANDIDATES = [
    DATA_DIR / "students.xlsx",
    DATA_DIR / "students.xls",
    BASE_DIR / "students.xlsx",
    BASE_DIR / "students.xls",
]

ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"
SLOTS_CACHE_FILE = DATA_DIR / "slots_cache.json"

LETTER_TEMPLATE = DATA_DIR / "letter_template.docx"

# أسماء الأعمدة الموجودة في ملفك (حسب الملف المرفق)
# ملاحظة: عمود "جهة التدريب " فيه مسافة في النهاية داخل الإكسل، سنقوم بتنظيفه تلقائيًا
COL_TRAINEE_ID = "رقم المتدرب"
COL_TRAINEE_NAME = "إسم المتدرب"
COL_PHONE = "رقم الجوال"
COL_PROGRAM = "برنامج"
COL_SPECIALTY = "التخصص"
COL_ENTITY = "جهة التدريب"
COL_DEPT = "القسم"
COL_TRAINER = "المدرب"
COL_COURSE = "اسم المقرر"
COL_REF = "الرقم المرجعي"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")


# -----------------------
# Helpers: JSON
# -----------------------
def _read_json(path: Path, default):
    try:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        pass
    return default


def _write_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


# -----------------------
# Helpers: Excel load
# -----------------------
def find_students_file() -> Path:
    for p in STUDENTS_CANDIDATES:
        if p.exists():
            return p
    raise FileNotFoundError(
        "لم يتم العثور على ملف الطلاب students.xlsx. "
        "ضعه داخل مجلد data باسم students.xlsx"
    )


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # تنظيف أسماء الأعمدة: إزالة مسافات البداية/النهاية وتوحيد بعض الفروقات
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # أحيانًا يظهر "اسم المتدرب" بدون همزة أو بصيغة مختلفة:
    # سنوحد إلى "إسم المتدرب" لو موجود بديل
    col_map = {}
    for c in df.columns:
        if c.strip() in ["اسم المتدرب", "إسم المتدرب", "اسم_المتدرب", "إسم_المتدرب"]:
            col_map[c] = COL_TRAINEE_NAME
        if c.strip() in ["جهة التدريب", "جهة التدريب "]:
            col_map[c] = COL_ENTITY

    if col_map:
        df = df.rename(columns=col_map)

    return df


def load_students_df() -> pd.DataFrame:
    students_file = find_students_file()
    df = pd.read_excel(students_file)
    df = normalize_columns(df)

    # تحقق من الأعمدة المطلوبة
    required = [
        COL_TRAINEE_ID, COL_TRAINEE_NAME, COL_PHONE,
        COL_PROGRAM, COL_SPECIALTY, COL_ENTITY
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(
            f"لم يتم العثور على الأعمدة المطلوبة: {missing} | "
            f"الأعمدة الموجودة: {list(df.columns)}"
        )

    # تنظيف القيم
    df[COL_TRAINEE_ID] = df[COL_TRAINEE_ID].astype(str).str.strip()
    df[COL_PHONE] = df[COL_PHONE].astype(str).str.strip()
    df[COL_PROGRAM] = df[COL_PROGRAM].astype(str).str.strip()
    df[COL_SPECIALTY] = df[COL_SPECIALTY].astype(str).str.strip()
    df[COL_ENTITY] = df[COL_ENTITY].astype(str).str.strip()
    df[COL_TRAINEE_NAME] = df[COL_TRAINEE_NAME].astype(str).str.strip()

    # قد يكون رقم الجوال بصيغة علمية/عشرية أحيانًا
    df[COL_PHONE] = df[COL_PHONE].str.replace(r"\.0$", "", regex=True)

    return df


def last4(phone: str) -> str:
    digits = "".join([ch for ch in str(phone) if ch.isdigit()])
    return digits[-4:] if len(digits) >= 4 else digits


# -----------------------
# فرص التدريب (Slots)
# -----------------------
def build_slots_from_excel(df: pd.DataFrame) -> dict:
    """
    يبني قاموس الفرص حسب (التخصص -> الجهة -> العدد)
    والعدد = عدد تكرار الجهة داخل تخصص/برنامج المتدرب في ملف الإكسل.
    """
    slots = {}
    for _, row in df.iterrows():
        spec = (row.get(COL_SPECIALTY, "") or "").strip()
        prog = (row.get(COL_PROGRAM, "") or "").strip()
        key = spec if spec else prog  # نعتمد التخصص أولاً، وإن كان فاضي نستخدم البرنامج
        entity = (row.get(COL_ENTITY, "") or "").strip()

        if not key or not entity:
            continue

        slots.setdefault(key, {})
        slots[key][entity] = slots[key].get(entity, 0) + 1

    return slots


def get_slots_cache(force_rebuild: bool = False) -> dict:
    df = load_students_df()

    if (not force_rebuild) and SLOTS_CACHE_FILE.exists():
        cached = _read_json(SLOTS_CACHE_FILE, {})
        if cached:
            return cached

    slots = build_slots_from_excel(df)
    _write_json(SLOTS_CACHE_FILE, slots)
    return slots


def get_assignments() -> dict:
    """
    assignments format:
    {
      "444229747": {
         "entity": "...",
         "specialty_key": "...",
         "ts": "2026-02-22T..."
      },
      ...
    }
    """
    return _read_json(ASSIGNMENTS_FILE, {})


def remaining_slots_for_key(slots: dict, specialty_key: str, assignments: dict) -> dict:
    """
    يرجّع المتبقي من الفرص للـ specialty_key بعد خصم اختيارات المتدربين.
    """
    base = dict(slots.get(specialty_key, {}))
    # خصم حسب assignments
    for tid, info in assignments.items():
        if info.get("specialty_key") != specialty_key:
            continue
        ent = info.get("entity")
        if ent in base:
            base[ent] = max(0, int(base[ent]) - 1)
    # حذف الجهات التي انتهت فرصها
    base = {k: v for k, v in base.items() if int(v) > 0}
    return base


# -----------------------
# DOCX placeholders replace
# -----------------------
def replace_in_docx(doc: Document, mapping: dict):
    # replace in paragraphs
    for p in doc.paragraphs:
        for k, v in mapping.items():
            if k in p.text:
                # بديل سريع: نجمع runs ونبدل ثم نعيد كتابته كنص واحد
                full = "".join(run.text for run in p.runs)
                full = full.replace(k, str(v))
                for run in p.runs:
                    run.text = ""
                if p.runs:
                    p.runs[0].text = full
                else:
                    p.add_run(full)

    # replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in mapping.items():
                        if k in p.text:
                            full = "".join(run.text for run in p.runs)
                            full = full.replace(k, str(v))
                            for run in p.runs:
                                run.text = ""
                            if p.runs:
                                p.runs[0].text = full
                            else:
                                p.add_run(full)


def convert_docx_to_pdf(docx_path: Path, out_dir: Path) -> Path:
    """
    يحاول التحويل عبر libreoffice (soffice).
    إذا لم يتوفر، يرفع خطأ واضح.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(docx_path),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_path = out_dir / (docx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise FileNotFoundError("فشل إنشاء PDF بعد التحويل.")
    return pdf_path


# -----------------------
# UI Templates
# -----------------------
BASE_CSS = """
<style>
  :root { --main:#0b1a34; --card:#ffffff; --bg:#f5f7fb; }
  body{
    margin:0; font-family: Arial, sans-serif; background: var(--bg);
    direction: rtl;
  }
  .top-image{
    width:100%;
    height:25vh;          /* ربع الصفحة تقريبًا */
    min-height:180px;
    max-height:320px;
    overflow:hidden;
    background:#eee;
  }
  .top-image img{
    width:100%;
    height:100%;
    object-fit:cover;      /* يمنع قصّ غريب */
    object-position:center;
    display:block;
  }
  .wrap{
    max-width:1100px;
    margin:-40px auto 60px auto;
    padding: 0 16px;
  }
  .card{
    background: var(--card);
    border-radius: 22px;
    box-shadow: 0 10px 30px rgba(0,0,0,.08);
    padding: 28px;
  }
  h1{
    margin:0 0 8px 0;
    font-size: 44px;
    text-align:center;
    font-weight: 800;
  }
  .sub{
    text-align:center;
    margin: 0 0 18px 0;
    color:#2b2b2b;
    font-size: 18px;
  }
  .row{
    display:flex;
    gap:16px;
    flex-wrap:wrap;
    justify-content:space-between;
    margin-top: 18px;
  }
  .field{
    flex: 1 1 300px;
  }
  label{
    display:block;
    font-weight:700;
    margin-bottom: 8px;
  }
  input, select{
    width:100%;
    padding:16px 14px;
    border-radius: 14px;
    border:1px solid #ddd;
    font-size: 18px;
    outline:none;
    background:#fff;
  }
  .btn{
    margin-top:18px;
    width:100%;
    padding:18px 16px;
    border:0;
    border-radius: 18px;
    background: var(--main);
    color:#fff;
    font-size: 20px;
    font-weight:700;
    cursor:pointer;
  }
  .error{
    color:#c40000;
    margin-top: 14px;
    font-weight:700;
    text-align:center;
  }
  .note{
    margin-top: 10px;
    color:#666;
    text-align:center;
  }
  .pill{
    display:inline-block;
    padding: 10px 14px;
    border-radius: 999px;
    background:#eef3ff;
    margin: 6px;
    font-weight:700;
  }
  .list{
    margin-top: 16px;
    line-height: 2;
  }
  .linkbtn{
    display:inline-block;
    margin-top: 10px;
    font-weight:800;
    text-decoration: underline;
  }
</style>
"""

LOGIN_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  """ + BASE_CSS + """
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>{{title}}</h1>
      <p class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

      <form method="post" action="{{ url_for('login') }}">
        <div class="row">
          <div class="field">
            <label>الرقم التدريبي</label>
            <input name="trainee_id" placeholder="مثال: 444229747" required>
          </div>
          <div class="field">
            <label>آخر 4 أرقام من الجوال</label>
            <input name="phone_last4" placeholder="مثال: 6101" required>
          </div>
        </div>

        <button class="btn" type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="error">{{error}}</div>
      {% endif %}

      <div class="note">ملاحظة: يتم قراءة بيانات الطلاب من <b>data/students.xlsx</b></div>
    </div>
  </div>
</body>
</html>
"""

CHOOSE_HTML = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  """ + BASE_CSS + """
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>{{title}}</h1>
      <p class="sub">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

      <div style="text-align:center; margin-top:10px;">
        <span class="pill">المتدرب: {{ trainee_name }}</span>
        <span class="pill">رقم المتدرب: {{ trainee_id }}</span>
        <span class="pill">التخصص/البرنامج: {{ specialty_key }}</span>
      </div>

      {% if already %}
        <hr style="margin:18px 0;">
        <div style="text-align:center; font-weight:800; font-size:18px;">
          تم تسجيل اختيارك مسبقًا:
        </div>
        <div style="text-align:center; margin-top:8px; font-size:18px;">
          الجهة المختارة: <b>{{ chosen_entity }}</b>
        </div>
        <div style="text-align:center;">
          <a class="linkbtn" href="{{ url_for('letter', tid=trainee_id) }}">تحميل/طباعة خطاب التوجيه PDF</a>
        </div>
      {% else %}
        <form method="post" action="{{ url_for('save_choice') }}">
          <input type="hidden" name="trainee_id" value="{{ trainee_id }}">

          <div class="row" style="margin-top:18px;">
            <div class="field">
              <label>جهة التدريب المتاحة</label>
              <select name="entity" required>
                <option value="">اختر الجهة...</option>
                {% for ent, cnt in options %}
                  <option value="{{ent}}">{{ent}} (متبقي: {{cnt}})</option>
                {% endfor %}
              </select>
            </div>
          </div>

          <button class="btn" type="submit">حفظ الاختيار</button>
        </form>

        {% if message %}
          <div class="error">{{message}}</div>
        {% endif %}
      {% endif %}

      <hr style="margin:18px 0;">

      <div style="font-weight:800; margin-bottom:8px;">ملخص الفرص المتبقية داخل تخصصك:</div>
      <div class="list">
        {% for ent, cnt in summary %}
          • {{ent}} : <b>{{cnt}}</b> فرصة<br>
        {% endfor %}
      </div>

    </div>
  </div>
</body>
</html>
"""


# -----------------------
# Routes
# -----------------------
@app.get("/")
def index():
    # نعرض نفس صفحة الدخول
    return render_template_string(LOGIN_HTML, title=APP_TITLE, error=None)


@app.get("/login")
def login_get():
    return redirect(url_for("index"))


@app.post("/")
def login():
    # نفس / (POST)
    trainee_id = (request.form.get("trainee_id") or "").strip()
    phone_last4 = (request.form.get("phone_last4") or "").strip()

    try:
        df = load_students_df()
    except Exception as e:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error=f"خطأ في قراءة ملف الإكسل: {e}")

    row = df[df[COL_TRAINEE_ID] == trainee_id]
    if row.empty:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error="الرقم التدريبي غير موجود.")

    r = row.iloc[0]
    if last4(r[COL_PHONE]) != phone_last4:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

    # نحدد مفتاح التخصص (التخصص ثم البرنامج)
    specialty_key = (str(r.get(COL_SPECIALTY, "")) or "").strip()
    if not specialty_key:
        specialty_key = (str(r.get(COL_PROGRAM, "")) or "").strip()

    trainee_name = str(r.get(COL_TRAINEE_NAME, "")).strip()

    # اذهب لصفحة الاختيار
    return redirect(url_for("choose", tid=trainee_id))


@app.get("/choose/<tid>")
def choose(tid):
    try:
        df = load_students_df()
    except Exception as e:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error=f"خطأ في قراءة ملف الإكسل: {e}")

    row = df[df[COL_TRAINEE_ID] == str(tid).strip()]
    if row.empty:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error="الرقم التدريبي غير موجود.")

    r = row.iloc[0]
    trainee_id = str(r[COL_TRAINEE_ID]).strip()
    trainee_name = str(r[COL_TRAINEE_NAME]).strip()
    specialty_key = (str(r.get(COL_SPECIALTY, "")) or "").strip()
    if not specialty_key:
        specialty_key = (str(r.get(COL_PROGRAM, "")) or "").strip()

    slots = get_slots_cache(force_rebuild=True)  # إعادة بناء للتأكد أنها مطابقة للإكسل الحالي
    assignments = get_assignments()

    # هل المتدرب اختار مسبقًا؟
    already = trainee_id in assignments
    chosen_entity = assignments.get(trainee_id, {}).get("entity") if already else None

    remaining = remaining_slots_for_key(slots, specialty_key, assignments)

    options = sorted(remaining.items(), key=lambda x: (-int(x[1]), x[0]))
    summary = options[:50]  # عرض أول 50 جهة (لو كثيرة)

    return render_template_string(
        CHOOSE_HTML,
        title=APP_TITLE,
        trainee_id=trainee_id,
        trainee_name=trainee_name,
        specialty_key=specialty_key,
        options=options,
        summary=summary,
        already=already,
        chosen_entity=chosen_entity,
        message=None,
    )


@app.post("/save-choice")
def save_choice():
    trainee_id = (request.form.get("trainee_id") or "").strip()
    entity = (request.form.get("entity") or "").strip()

    try:
        df = load_students_df()
    except Exception as e:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error=f"خطأ في قراءة ملف الإكسل: {e}")

    row = df[df[COL_TRAINEE_ID] == trainee_id]
    if row.empty:
        return render_template_string(LOGIN_HTML, title=APP_TITLE, error="الرقم التدريبي غير موجود.")

    r = row.iloc[0]
    specialty_key = (str(r.get(COL_SPECIALTY, "")) or "").strip()
    if not specialty_key:
        specialty_key = (str(r.get(COL_PROGRAM, "")) or "").strip()

    slots = get_slots_cache(force_rebuild=True)
    assignments = get_assignments()

    # منع إعادة الاختيار لو اختار مسبقًا
    if trainee_id in assignments:
        return redirect(url_for("choose", tid=trainee_id))

    remaining = remaining_slots_for_key(slots, specialty_key, assignments)

    if entity not in remaining:
        # لا توجد فرص متاحة
        return render_template_string(
            CHOOSE_HTML,
            title=APP_TITLE,
            trainee_id=trainee_id,
            trainee_name=str(r.get(COL_TRAINEE_NAME, "")).strip(),
            specialty_key=specialty_key,
            options=sorted(remaining.items(), key=lambda x: (-int(x[1]), x[0])),
            summary=sorted(remaining.items(), key=lambda x: (-int(x[1]), x[0]))[:50],
            already=False,
            chosen_entity=None,
            message="لا توجد فرص متاحة لهذه الجهة (قد تكون انتهت). اختر جهة أخرى.",
        )

    # حفظ الاختيار
    assignments[trainee_id] = {
        "entity": entity,
        "specialty_key": specialty_key,
        "ts": datetime.utcnow().isoformat(),
    }
    _write_json(ASSIGNMENTS_FILE, assignments)

    return redirect(url_for("choose", tid=trainee_id))


@app.get("/letter/<tid>")
def letter(tid):
    """
    توليد خطاب التوجيه:
    - ينشئ DOCX من template
    - يحوّله PDF عبر soffice (LibreOffice)
    - إذا فشل التحويل، يرجّع DOCX بدل PDF مع رسالة خطأ واضحة
    """
    trainee_id = str(tid).strip()

    if not LETTER_TEMPLATE.exists():
        abort(500, "ملف القالب غير موجود: data/letter_template.docx")

    try:
        df = load_students_df()
    except Exception as e:
        abort(500, f"خطأ في قراءة ملف الإكسل: {e}")

    row = df[df[COL_TRAINEE_ID] == trainee_id]
    if row.empty:
        abort(404, "المتدرب غير موجود.")

    assignments = get_assignments()
    if trainee_id not in assignments:
        abort(400, "لم يتم اختيار جهة تدريب بعد.")

    entity = assignments[trainee_id]["entity"]

    r = row.iloc[0]
    mapping = {
        "{{trainee_name}}": str(r.get(COL_TRAINEE_NAME, "")).strip(),
        "{{trainee_id}}": trainee_id,
        "{{phone}}": str(r.get(COL_PHONE, "")).strip(),
        "{{program}}": str(r.get(COL_PROGRAM, "")).strip(),
        "{{specialty}}": str(r.get(COL_SPECIALTY, "")).strip(),
        "{{department}}": str(r.get(COL_DEPT, "")).strip(),
        "{{trainer}}": str(r.get(COL_TRAINER, "")).strip(),
        "{{course_name}}": str(r.get(COL_COURSE, "")).strip(),
        "{{reference_no}}": str(r.get(COL_REF, "")).strip(),
        "{{entity}}": entity,
        "{{date}}": datetime.now().strftime("%Y-%m-%d"),
    }

    out_dir = DATA_DIR / "out"
    out_dir.mkdir(parents=True, exist_ok=True)

    doc = Document(str(LETTER_TEMPLATE))
    replace_in_docx(doc, mapping)

    docx_path = out_dir / f"letter_{trainee_id}.docx"
    doc.save(str(docx_path))

    # حاول PDF
    try:
        pdf_path = convert_docx_to_pdf(docx_path, out_dir)
        return send_file(pdf_path, as_attachment=True, download_name=f"letter_{trainee_id}.pdf")
    except Exception as e:
        # رجّع docx إذا PDF فشل (عادة بسبب soffice)
        return send_file(
            docx_path,
            as_attachment=True,
            download_name=f"letter_{trainee_id}.docx",
        )


# دعم /letter?tid=...
@app.get("/letter")
def letter_query():
    tid = request.args.get("tid")
    if not tid:
        abort(400, "ضع tid")
    return redirect(url_for("letter", tid=str(tid).strip()))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

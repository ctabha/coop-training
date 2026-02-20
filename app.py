import os
import json
import re
import uuid
import shutil
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, send_file, redirect, url_for, render_template_string

# -----------------------------
# App Config
# -----------------------------
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

STUDENTS_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"
SLOTS_FILE = DATA_DIR / "slots.json"
TEMPLATE_DOCX = DATA_DIR / "letter_template.docx"

# Excel column names (after we strip whitespace)
COL_TRAINEE_ID = "رقم المتدرب"
COL_PHONE = "رقم الجوال"
COL_TRAINEE_NAME = "اسم المتدرب"
COL_SPECIALTY = "التخصص"
COL_ENTITY = "جهة التدريب"
COL_PROGRAM = "برنامج"

# -----------------------------
# Helpers
# -----------------------------
def clean_text(x) -> str:
    """Normalize Arabic/English strings, remove NBSP and extra spaces."""
    if x is None:
        return ""
    s = str(x)
    s = s.replace("\u00a0", " ").strip()
    s = re.sub(r"\s+", " ", s)
    if s.lower() in ("nan", "none", "null"):
        return ""
    return s


def load_students() -> pd.DataFrame:
    if not STUDENTS_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {STUDENTS_FILE}")

    df = pd.read_excel(STUDENTS_FILE, engine="openpyxl")

    # IMPORTANT: remove trailing/leading spaces from column headers
    df.columns = [clean_text(c) for c in df.columns]

    # Validate required columns
    required = [COL_TRAINEE_ID, COL_PHONE, COL_TRAINEE_NAME, COL_SPECIALTY, COL_ENTITY]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"لم أجد الأعمدة المطلوبة: {missing}. "
            f"الأعمدة الموجودة: {list(df.columns)}"
        )

    # Normalize key columns
    for c in [COL_TRAINEE_ID, COL_PHONE, COL_TRAINEE_NAME, COL_SPECIALTY, COL_ENTITY]:
        df[c] = df[c].apply(clean_text)

    # Ensure trainee id is string
    df[COL_TRAINEE_ID] = df[COL_TRAINEE_ID].astype(str).apply(clean_text)

    # phone may be numeric -> keep digits
    def phone_digits(v):
        s = clean_text(v)
        s = re.sub(r"\D+", "", s)  # only digits
        return s

    df[COL_PHONE] = df[COL_PHONE].apply(phone_digits)

    # Optional columns
    if COL_PROGRAM in df.columns:
        df[COL_PROGRAM] = df[COL_PROGRAM].apply(clean_text)
    else:
        df[COL_PROGRAM] = ""

    return df


def load_json(path: Path, default):
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return default


def save_json(path: Path, data):
    tmp = str(path) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def compute_initial_slots(df: pd.DataFrame) -> dict:
    """
    slots[specialty][entity] = count of rows where (specialty, entity) appears.
    """
    x = df.copy()
    x[COL_SPECIALTY] = x[COL_SPECIALTY].apply(clean_text)
    x[COL_ENTITY] = x[COL_ENTITY].apply(clean_text)

    # drop empty entity/specialty
    x = x[(x[COL_SPECIALTY] != "") & (x[COL_ENTITY] != "")]

    g = x.groupby([COL_SPECIALTY, COL_ENTITY]).size().reset_index(name="count")

    slots = {}
    for _, r in g.iterrows():
        spec = r[COL_SPECIALTY]
        ent = r[COL_ENTITY]
        cnt = int(r["count"])
        slots.setdefault(spec, {})
        slots[spec][ent] = cnt

    return slots


def ensure_slots(df: pd.DataFrame) -> dict:
    """
    Create slots.json if missing, else load it.
    """
    slots = load_json(SLOTS_FILE, default=None)
    if not slots:
        slots = compute_initial_slots(df)
        save_json(SLOTS_FILE, slots)
    return slots


def last4(phone_digits_str: str) -> str:
    s = re.sub(r"\D+", "", phone_digits_str or "")
    return s[-4:] if len(s) >= 4 else s


def find_student(df: pd.DataFrame, trainee_id: str, phone_last4: str):
    trainee_id = clean_text(trainee_id)
    phone_last4 = clean_text(phone_last4)

    row = df[df[COL_TRAINEE_ID] == trainee_id]
    if row.empty:
        return None

    r = row.iloc[0].to_dict()
    if last4(r.get(COL_PHONE, "")) != phone_last4:
        return "WRONG_PHONE"

    return r


def libreoffice_convert_to_pdf(input_docx: Path, out_dir: Path) -> Path:
    """
    Convert DOCX to PDF using LibreOffice (installed in Dockerfile).
    """
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_dir),
        str(input_docx),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    pdf_path = out_dir / (input_docx.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError("فشل تحويل الملف إلى PDF.")
    return pdf_path


def fill_docx_template_copy(template_path: Path, out_docx_path: Path, replacements: dict):
    """
    Minimal DOCX replace by using python-docx (simple placeholders).
    Placeholders in template like: {{TRAINEE_NAME}}, {{TRAINEE_ID}}, {{ENTITY}}, {{DATE}}
    """
    from docx import Document

    doc = Document(str(template_path))

    def replace_in_paragraph(paragraph):
        for key, val in replacements.items():
            if key in paragraph.text:
                # Replace across runs (best-effort)
                for run in paragraph.runs:
                    run.text = run.text.replace(key, val)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)

    doc.save(str(out_docx_path))


# -----------------------------
# Flask App
# -----------------------------
app = Flask(__name__)


PAGE_LOGIN = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
  body{font-family:Arial;background:#f7f7f7;margin:0}
  .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
  .top-image img{width:100%;height:100%;object-fit:cover;object-position:center;display:block}
  .wrap{max-width:980px;margin:20px auto;padding:0 16px}
  .card{background:#fff;border-radius:18px;box-shadow:0 8px 24px rgba(0,0,0,.08);padding:28px}
  h1{margin:0 0 8px 0;font-size:44px;text-align:center}
  p{margin:0 0 20px 0;text-align:center;color:#444}
  .row{display:flex;gap:18px;flex-wrap:wrap}
  .field{flex:1;min-width:260px}
  label{display:block;margin:0 0 10px 0;font-weight:700}
  input{width:100%;padding:16px;border-radius:14px;border:1px solid #ddd;font-size:20px}
  button{margin-top:18px;width:100%;padding:18px;border:0;border-radius:18px;background:#0b1630;color:#fff;font-size:22px;cursor:pointer}
  .err{margin-top:14px;color:#c00;font-weight:700;text-align:center}
  .note{margin-top:10px;color:#666;text-align:center}
</style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

      <form method="post" action="/">
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
        <button type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}

      {% if sys_error %}
        <div class="err">حدث خطأ أثناء التحميل/التحقق: {{sys_error}}</div>
      {% endif %}

      <div class="note">ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b></div>
    </div>
  </div>
</body>
</html>
"""


PAGE_PICK = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
  body{font-family:Arial;background:#f7f7f7;margin:0}
  .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
  .top-image img{width:100%;height:100%;object-fit:cover;object-position:center;display:block}
  .wrap{max-width:980px;margin:20px auto;padding:0 16px}
  .card{background:#fff;border-radius:18px;box-shadow:0 8px 24px rgba(0,0,0,.08);padding:28px}
  h1{margin:0 0 6px 0;font-size:38px;text-align:center}
  .sub{margin:0 0 20px 0;text-align:center;color:#444}
  .pillrow{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;margin:10px 0 20px}
  .pill{background:#eef3ff;padding:10px 14px;border-radius:999px;font-weight:700}
  label{display:block;margin:14px 0 10px 0;font-weight:700;font-size:20px}
  select{width:100%;padding:16px;border-radius:14px;border:1px solid #ddd;font-size:20px;background:#fff}
  button{margin-top:18px;width:100%;padding:18px;border:0;border-radius:18px;background:#0b1630;color:#fff;font-size:22px;cursor:pointer}
  .err{margin-top:14px;color:#c00;font-weight:700;text-align:center}
  .ok{margin-top:14px;color:#0a7a2a;font-weight:700;text-align:center}
  .row2{display:flex;gap:12px;margin-top:14px}
  .row2 a{flex:1;text-decoration:none;display:block;text-align:center;padding:14px;border-radius:14px;background:#111827;color:#fff}
  .row2 a.secondary{background:#334155}
</style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p class="sub">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

      <div class="pillrow">
        <div class="pill">المتدرب: {{trainee_name}}</div>
        <div class="pill">رقم المتدرب: {{trainee_id}}</div>
        <div class="pill">التخصص: {{specialty}}</div>
      </div>

      <form method="post" action="/pick">
        <input type="hidden" name="trainee_id" value="{{trainee_id}}">
        <input type="hidden" name="phone_last4" value="{{phone_last4}}">

        <label>جهة التدريب المتاحة</label>
        <select name="entity" required>
          <option value="">اختر الجهة...</option>
          {% for ent, remaining in options %}
            <option value="{{ent}}">{{ent}} — ({{remaining}} فرصة)</option>
          {% endfor %}
        </select>

        <button type="submit">حفظ الاختيار</button>
      </form>

      {% if ok %}
        <div class="ok">{{ok}}</div>
        <div class="row2">
          <a href="/print?trainee_id={{trainee_id}}&phone_last4={{phone_last4}}">طباعة خطاب التوجيه PDF</a>
          <a class="secondary" href="/">تسجيل خروج</a>
        </div>
      {% endif %}

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}
    </div>
  </div>
</body>
</html>
"""


def get_remaining_options(slots: dict, specialty: str) -> list:
    """
    returns list of (entity, remaining) for that specialty, remaining>0
    """
    specialty = clean_text(specialty)
    d = slots.get(specialty, {}) if isinstance(slots, dict) else {}
    opts = []
    for ent, cnt in d.items():
        try:
            c = int(cnt)
        except Exception:
            c = 0
        if c > 0:
            opts.append((ent, c))
    # Sort by remaining desc then name
    opts.sort(key=lambda x: (-x[1], x[0]))
    return opts


@app.route("/", methods=["GET", "POST"])
def home():
    error = ""
    sys_error = ""
    if request.method == "POST":
        trainee_id = request.form.get("trainee_id", "")
        phone_last4 = request.form.get("phone_last4", "")

        try:
            df = load_students()
            _ = ensure_slots(df)
            st = find_student(df, trainee_id, phone_last4)

            if st is None:
                error = "الرقم التدريبي غير موجود."
            elif st == "WRONG_PHONE":
                error = "بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال."
            else:
                # success -> go pick
                return redirect(url_for("pick", trainee_id=trainee_id, phone_last4=phone_last4))
        except Exception as e:
            sys_error = str(e)

    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error, sys_error=sys_error)


@app.route("/pick", methods=["GET", "POST"])
def pick():
    try:
        df = load_students()
        slots = ensure_slots(df)
        assignments = load_json(ASSIGNMENTS_FILE, default={})
    except Exception as e:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="", sys_error=str(e))

    if request.method == "GET":
        trainee_id = request.args.get("trainee_id", "")
        phone_last4 = request.args.get("phone_last4", "")
    else:
        trainee_id = request.form.get("trainee_id", "")
        phone_last4 = request.form.get("phone_last4", "")

    st = find_student(df, trainee_id, phone_last4)
    if st is None:
        return redirect(url_for("home"))
    if st == "WRONG_PHONE":
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="بيانات الدخول غير صحيحة.", sys_error="")

    trainee_name = st.get(COL_TRAINEE_NAME, "")
    specialty = st.get(COL_SPECIALTY, "")

    ok = ""
    error = ""

    # If already assigned, show message
    already = assignments.get(str(trainee_id))
    if request.method == "POST":
        chosen = clean_text(request.form.get("entity", ""))
        if not chosen:
            error = "اختر جهة تدريب."
        else:
            # Ensure still available
            remaining = int(slots.get(specialty, {}).get(chosen, 0) or 0)
            if remaining <= 0:
                error = "عذرًا، هذه الجهة لم تعد متاحة."
            else:
                # If trainee already had an assignment, return it back first (so they can change)
                if already and already.get("entity") and already.get("specialty") == specialty:
                    prev_ent = already.get("entity")
                    if prev_ent in slots.get(specialty, {}):
                        slots[specialty][prev_ent] = int(slots[specialty][prev_ent]) + 1

                # Decrement selected
                slots.setdefault(specialty, {})
                slots[specialty][chosen] = int(slots[specialty].get(chosen, 0)) - 1

                # Save assignment
                assignments[str(trainee_id)] = {
                    "trainee_id": str(trainee_id),
                    "trainee_name": trainee_name,
                    "specialty": specialty,
                    "entity": chosen,
                    "saved_at": datetime.utcnow().isoformat() + "Z",
                }

                save_json(SLOTS_FILE, slots)
                save_json(ASSIGNMENTS_FILE, assignments)

                ok = "تم حفظ اختيار جهة التدريب بنجاح."

    # reload after update
    slots = load_json(SLOTS_FILE, default={})
    assignments = load_json(ASSIGNMENTS_FILE, default={})
    already = assignments.get(str(trainee_id))

    options = get_remaining_options(slots, specialty)

    # If already assigned, keep it visible even لو صارت 0 (اختياري)
    # هنا نخليها تظهر فقط ضمن الخيارات المتاحة.

    if not options:
        error = "لا توجد جهات متاحة حاليًا لهذا التخصص."

    return render_template_string(
        PAGE_PICK,
        title=APP_TITLE,
        trainee_id=str(trainee_id),
        phone_last4=phone_last4,
        trainee_name=trainee_name,
        specialty=specialty,
        options=options,
        ok=ok,
        error=error,
    )


@app.route("/print", methods=["GET"])
def print_letter():
    try:
        df = load_students()
        assignments = load_json(ASSIGNMENTS_FILE, default={})
    except Exception as e:
        return f"خطأ: {e}", 500

    trainee_id = request.args.get("trainee_id", "")
    phone_last4 = request.args.get("phone_last4", "")

    st = find_student(df, trainee_id, phone_last4)
    if st is None or st == "WRONG_PHONE":
        return redirect(url_for("home"))

    ass = assignments.get(str(trainee_id))
    if not ass:
        return "لم يتم اختيار جهة تدريب بعد. ارجع واختر جهة ثم اطبع.", 400

    if not TEMPLATE_DOCX.exists():
        return f"ملف القالب غير موجود: {TEMPLATE_DOCX}", 500

    # Prepare output files
    job_id = uuid.uuid4().hex[:10]
    out_docx = OUT_DIR / f"letter_{trainee_id}_{job_id}.docx"

    replacements = {
        "{{TRAINEE_NAME}}": clean_text(ass.get("trainee_name", "")),
        "{{TRAINEE_ID}}": clean_text(ass.get("trainee_id", "")),
        "{{SPECIALTY}}": clean_text(ass.get("specialty", "")),
        "{{ENTITY}}": clean_text(ass.get("entity", "")),
        "{{DATE}}": datetime.now().strftime("%Y/%m/%d"),
    }

    try:
        fill_docx_template_copy(TEMPLATE_DOCX, out_docx, replacements)
        pdf_path = libreoffice_convert_to_pdf(out_docx, OUT_DIR)
        return send_file(pdf_path, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")
    except subprocess.CalledProcessError as e:
        return f"فشل التحويل إلى PDF. تفاصيل: {e}", 500
    except Exception as e:
        return f"خطأ أثناء الطباعة: {e}", 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

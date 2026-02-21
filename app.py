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
    redirect,
    url_for,
    send_file,
    render_template_string,
    session,
)

# =========================
# إعدادات عامة
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

DATA_FILE = DATA_DIR / "students.xlsx"
TEMPLATE_FILE = DATA_DIR / "letter_template.docx"

ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"  # اختيارات المتدربين

# =========================
# Flask App
# =========================
app = Flask(__name__, static_folder=str(STATIC_DIR))
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-key")


# =========================
# أدوات مساعدة
# =========================
def _safe_mkdir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _load_json(path: Path, default):
    try:
        if path.exists():
            with path.open("r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default


def _save_json_atomic(path: Path, obj) -> None:
    _safe_mkdir(path.parent)
    tmp = path.with_suffix(".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)
    tmp.replace(path)


def normalize_digits(s: str) -> str:
    """يحافظ على الأرقام فقط."""
    if s is None:
        return ""
    s = str(s).strip()
    return re.sub(r"\D+", "", s)


def last4_phone(v) -> str:
    d = normalize_digits(v)
    return d[-4:] if len(d) >= 4 else d


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    # إزالة الفراغات من أسماء الأعمدة (مهم لأن عندك "جهة التدريب " فيها مسافة)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def required_columns_present(df: pd.DataFrame) -> (bool, str):
    required = [
        "رقم المتدرب",
        "رقم الجوال",
        "إسم المتدرب",
        "التخصص",
        "برنامج",
        "جهة التدريب",
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return False, f"الأعمدة الناقصة في ملف الإكسل: {missing}\nالأعمدة الموجودة: {list(df.columns)}"
    return True, ""


def load_students() -> pd.DataFrame:
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE)  # يحتاج openpyxl في requirements
    df = normalize_columns(df)

    ok, msg = required_columns_present(df)
    if not ok:
        raise ValueError(msg)

    # تجهيز حقول مطابقة الدخول
    df["رقم المتدرب"] = df["رقم المتدرب"].apply(normalize_digits)
    df["رقم الجوال"] = df["رقم الجوال"].apply(normalize_digits)
    df["اخر4"] = df["رقم الجوال"].apply(lambda x: x[-4:] if len(x) >= 4 else x)

    # تنظيف جهة التدريب
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()

    return df


def capacity_by_specialty(df: pd.DataFrame) -> dict:
    """
    يحسب الفرص من ملف الاكسل:
    كل جهة تدريبية مكررة داخل نفس التخصص = عدد فرص.
    """
    grp = (
        df.groupby(["التخصص", "جهة التدريب"])
        .size()
        .reset_index(name="count")
    )
    cap = {}
    for _, row in grp.iterrows():
        spec = str(row["التخصص"]).strip()
        ent = str(row["جهة التدريب"]).strip()
        cap.setdefault(spec, {})[ent] = int(row["count"])
    return cap


def load_assignments() -> dict:
    return _load_json(ASSIGNMENTS_FILE, {})


def remaining_options_for_student(df: pd.DataFrame, trainee_id: str) -> dict:
    """
    يرجع:
    - student row
    - options list (جهات متاحة ضمن تخصصه مع العدد المتبقي)
    - remaining_map
    """
    trainee_id = normalize_digits(trainee_id)
    row = df[df["رقم المتدرب"] == trainee_id]
    if row.empty:
        return None, [], {}

    student = row.iloc[0].to_dict()
    spec = str(student["التخصص"]).strip()

    cap = capacity_by_specialty(df)
    spec_cap = cap.get(spec, {})

    assignments = load_assignments()

    # حساب الاستهلاك داخل نفس التخصص
    used = {}
    for tid, rec in assignments.items():
        try:
            if str(rec.get("التخصص", "")).strip() != spec:
                continue
            ent = str(rec.get("جهة التدريب", "")).strip()
            if ent:
                used[ent] = used.get(ent, 0) + 1
        except Exception:
            continue

    remaining = {}
    for ent, cnt in spec_cap.items():
        remaining[ent] = max(0, int(cnt) - int(used.get(ent, 0)))

    # الخيارات المتاحة (المتبقي > 0)
    options = [(ent, remaining[ent]) for ent in remaining if remaining[ent] > 0]
    options.sort(key=lambda x: (-x[1], x[0]))

    return student, options, remaining


def ensure_logged_in():
    tid = session.get("tid")
    if not tid:
        return None
    return normalize_digits(tid)


# =========================
# HTML (واجهة)
# =========================
BASE_CSS = """
<style>
  body{
    font-family: Arial, sans-serif;
    background:#f7f7f7;
    margin:0;
    direction:rtl;
  }
  .top-image{
    width:100%;
    height:25vh;            /* ربع الصفحة */
    overflow:hidden;
    background:#fff;
  }
  .top-image img{
    width:100%;
    height:100%;
    object-fit:contain;     /* يمنع القص */
    display:block;
  }
  .wrap{
    max-width:1000px;
    margin: 24px auto;
    padding: 0 16px;
  }
  .card{
    background:#fff;
    border-radius:18px;
    box-shadow: 0 8px 25px rgba(0,0,0,.08);
    padding: 28px;
  }
  h1{
    margin:0 0 8px 0;
    font-size:44px;
    text-align:center;
    font-weight:800;
  }
  .subtitle{
    text-align:center;
    color:#555;
    margin-bottom:18px;
  }
  .row{
    display:flex;
    gap:16px;
    flex-wrap:wrap;
    margin-top:18px;
  }
  .col{
    flex: 1 1 320px;
  }
  label{
    font-weight:700;
    display:block;
    margin-bottom:8px;
  }
  input, select{
    width:100%;
    padding:14px 16px;
    border-radius:14px;
    border:1px solid #ddd;
    font-size:18px;
    outline:none;
  }
  .btn{
    width:100%;
    margin-top:16px;
    padding:18px 16px;
    border-radius:18px;
    border:none;
    background:#0b1630;
    color:#fff;
    font-size:20px;
    cursor:pointer;
    font-weight:700;
  }
  .msg{
    margin-top:14px;
    text-align:center;
  }
  .err{ color:#c00; font-weight:700; }
  .ok{ color:#0a7; font-weight:700; }
  .pill{
    display:inline-block;
    background:#eef3ff;
    padding:10px 14px;
    border-radius:999px;
    font-weight:700;
    margin: 6px 6px;
  }
  .small{
    color:#666;
    font-size:14px;
    margin-top:8px;
  }
  a{
    color:#0b1630;
    font-weight:700;
  }
</style>
"""

LOGIN_PAGE = """
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
      <div class="subtitle">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</div>

      <form method="POST" action="/">
        <div class="row">
          <div class="col">
            <label>الرقم التدريبي</label>
            <input name="tid" placeholder="مثال: 444229747" required>
          </div>
          <div class="col">
            <label>آخر 4 أرقام من الجوال</label>
            <input name="last4" placeholder="مثال: 6101" required>
          </div>
        </div>

        <button class="btn" type="submit">دخول</button>

        {% if msg %}
          <div class="msg {{'err' if is_err else 'ok'}}">{{msg}}</div>
        {% endif %}

        <div class="small">
          ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b>
        </div>
      </form>
    </div>
  </div>
</body>
</html>
"""

SELECT_PAGE = """
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
      <div class="subtitle">اختر جهة التدريب المتاحة لتخصصك ثم اطبع خطاب التوجيه.</div>

      <div style="text-align:center;margin-top:8px">
        <span class="pill">المتدرب: {{name}}</span>
        <span class="pill">رقم المتدرب: {{tid}}</span>
        <span class="pill">التخصص: {{spec}}</span>
        <span class="pill">البرنامج: {{program}}</span>
      </div>

      <hr style="margin:18px 0;border:none;border-top:1px solid #eee;">

      {% if already %}
        <div class="msg ok">
          تم تسجيل اختيارك مسبقًا:
          <b>{{chosen}}</b>
        </div>

        <div class="msg" style="margin-top:10px">
          <a href="/letter">تحميل/طباعة خطاب التوجيه (PDF أو Word)</a>
        </div>

      {% else %}
        <form method="POST" action="/select">
          <div class="row">
            <div class="col">
              <label>جهة التدريب المتاحة (حسب تخصصك)</label>
              <select name="entity" required>
                <option value="">اختر الجهة...</option>
                {% for ent, rem in options %}
                  <option value="{{ent}}">{{ent}} (متبقي: {{rem}})</option>
                {% endfor %}
              </select>
            </div>
          </div>

          <button class="btn" type="submit">حفظ الاختيار</button>

          {% if msg %}
            <div class="msg {{'err' if is_err else 'ok'}}">{{msg}}</div>
          {% endif %}
        </form>

        <div class="small" style="text-align:center;margin-top:10px">
          ملاحظة: يتم احتساب الفرص من تكرار <b>جهة التدريب</b> داخل نفس <b>التخصص</b> في ملف الإكسل،
          وتتناقص تلقائيًا مع كل اختيار.
        </div>
      {% endif %}

      {% if remaining_summary %}
        <hr style="margin:18px 0;border:none;border-top:1px solid #eee;">
        <div style="font-weight:800;margin-bottom:8px">ملخص الفرص المتبقية داخل تخصصك:</div>
        <ul>
          {% for ent, rem in remaining_summary %}
            <li>{{ent}} : {{rem}} فرصة</li>
          {% endfor %}
        </ul>
      {% endif %}
    </div>
  </div>
</body>
</html>
"""


# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
def home():
    msg = ""
    is_err = False

    try:
        df = load_students()
    except Exception as e:
        return render_template_string(
            LOGIN_PAGE,
            title=APP_TITLE,
            msg=f"حدث خطأ أثناء التحميل/التحقق: {e}",
            is_err=True
        )

    if request.method == "POST":
        tid = normalize_digits(request.form.get("tid", ""))
        last4 = normalize_digits(request.form.get("last4", ""))

        row = df[df["رقم المتدرب"] == tid]
        if row.empty:
            msg = "بيانات الدخول غير صحيحة: الرقم التدريبي غير موجود."
            is_err = True
        else:
            phone_last4 = str(row.iloc[0]["اخر4"])
            if phone_last4 != last4:
                msg = "بيانات الدخول غير صحيحة: تأكد من آخر 4 أرقام من الجوال."
                is_err = True
            else:
                session["tid"] = tid
                return redirect(url_for("select"))

    return render_template_string(LOGIN_PAGE, title=APP_TITLE, msg=msg, is_err=is_err)


@app.route("/select", methods=["GET", "POST"])
def select():
    tid = ensure_logged_in()
    if not tid:
        return redirect(url_for("home"))

    try:
        df = load_students()
    except Exception as e:
        return f"خطأ في قراءة ملف الطلاب: {e}", 500

    assignments = load_assignments()
    if tid in assignments:
        rec = assignments[tid]
        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            name=rec.get("إسم المتدرب", ""),
            tid=tid,
            spec=rec.get("التخصص", ""),
            program=rec.get("برنامج", ""),
            already=True,
            chosen=rec.get("جهة التدريب", ""),
            options=[],
            msg="",
            is_err=False,
            remaining_summary=[]
        )

    student, options, remaining = remaining_options_for_student(df, tid)
    if not student:
        return redirect(url_for("home"))

    msg = ""
    is_err = False

    if request.method == "POST":
        chosen = (request.form.get("entity") or "").strip()
        if not chosen:
            msg = "اختر جهة تدريب أولاً."
            is_err = True
        else:
            # إعادة فحص التوفر لحظة الحفظ (حماية من تعارض)
            _, options_now, remaining_now = remaining_options_for_student(df, tid)
            rem = remaining_now.get(chosen, 0)
            if rem <= 0:
                msg = "عذرًا، هذه الجهة لم تعد متاحة الآن. اختر جهة أخرى."
                is_err = True
                options = options_now
                remaining = remaining_now
            else:
                assignments[tid] = {
                    "رقم المتدرب": tid,
                    "إسم المتدرب": str(student.get("إسم المتدرب", "")).strip(),
                    "التخصص": str(student.get("التخصص", "")).strip(),
                    "برنامج": str(student.get("برنامج", "")).strip(),
                    "رقم الجوال": str(student.get("رقم الجوال", "")).strip(),
                    "جهة التدريب": chosen,
                    "ts": datetime.utcnow().isoformat()
                }
                _save_json_atomic(ASSIGNMENTS_FILE, assignments)
                return redirect(url_for("select"))

    # ملخص متبقي (داخل التخصص فقط)
    remaining_summary = [(k, v) for k, v in remaining.items()]
    remaining_summary.sort(key=lambda x: (-x[1], x[0]))

    return render_template_string(
        SELECT_PAGE,
        title=APP_TITLE,
        name=str(student.get("إسم المتدرب", "")).strip(),
        tid=tid,
        spec=str(student.get("التخصص", "")).strip(),
        program=str(student.get("برنامج", "")).strip(),
        already=False,
        chosen="",
        options=options,
        msg=msg,
        is_err=is_err,
        remaining_summary=remaining_summary[:30],  # اختصار العرض
    )


def _replace_text_in_docx(doc, mapping: dict):
    """
    استبدال نصوص بسيطة داخل docx (paragraphs + tables).
    ضع في القالب نصوص مثل: {{NAME}} {{ENTITY}} ... إلخ
    """
    # الفقرات
    for p in doc.paragraphs:
        for k, v in mapping.items():
            if k in p.text:
                for r in p.runs:
                    r.text = r.text.replace(k, v)

    # الجداول
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in mapping.items():
                        if k in p.text:
                            for r in p.runs:
                                r.text = r.text.replace(k, v)


def _find_office_exe():
    for exe in ["soffice", "libreoffice"]:
        if shutil.which(exe):
            return exe
    return None


@app.route("/letter", methods=["GET"])
def letter():
    tid = ensure_logged_in()
    if not tid:
        return redirect(url_for("home"))

    assignments = load_assignments()
    if tid not in assignments:
        return redirect(url_for("select"))

    rec = assignments[tid]

    if not TEMPLATE_FILE.exists():
        return f"قالب الخطاب غير موجود: {TEMPLATE_FILE}", 500

    # إنشاء DOCX مخصص
    from docx import Document  # python-docx

    tmp_dir = Path("/tmp") / f"letter_{uuid.uuid4().hex}"
    _safe_mkdir(tmp_dir)

    out_docx = tmp_dir / f"letter_{tid}.docx"
    out_pdf = tmp_dir / f"letter_{tid}.pdf"

    doc = Document(str(TEMPLATE_FILE))

    mapping = {
        "{{NAME}}": str(rec.get("إسم المتدرب", "")).strip(),
        "{{TID}}": str(rec.get("رقم المتدرب", "")).strip(),
        "{{SPEC}}": str(rec.get("التخصص", "")).strip(),
        "{{PROGRAM}}": str(rec.get("برنامج", "")).strip(),
        "{{ENTITY}}": str(rec.get("جهة التدريب", "")).strip(),
        "{{DATE}}": datetime.now().strftime("%Y-%m-%d"),
    }
    _replace_text_in_docx(doc, mapping)
    doc.save(str(out_docx))

    # محاولة تحويل PDF إن توفر LibreOffice
    office = _find_office_exe()
    if office:
        try:
            subprocess.run(
                [
                    office,
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(tmp_dir),
                    str(out_docx),
                ],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=120,
            )
            if out_pdf.exists():
                return send_file(
                    str(out_pdf),
                    as_attachment=True,
                    download_name=f"letter_{tid}.pdf",
                    mimetype="application/pdf",
                )
        except Exception:
            # لو فشل التحويل لأي سبب، نكمل ونرسل DOCX
            pass

    # fallback: ارسال Word
    return send_file(
        str(out_docx),
        as_attachment=True,
        download_name=f"letter_{tid}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

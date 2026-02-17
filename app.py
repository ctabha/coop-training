import os
import re
import uuid
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, session, send_file, abort

from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

# أقصى عدد جهات تظهر للطالب (اختياري)
MAX_SLOTS_PER_ENTITY = 5

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "out"
OUT_DIR.mkdir(exist_ok=True)

# ملفات البيانات (تقدر تغير أسماءها)
STUDENTS_XLSX = DATA_DIR / "students.xlsx"
STUDENTS_CSV = DATA_DIR / "students.csv"

ENTITIES_XLSX = DATA_DIR / "entities.xlsx"
ENTITIES_CSV = DATA_DIR / "entities.csv"

# صورة الهيدر (لازم تكون داخل static)
HEADER_IMG = "header.jpg"

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me-please")


# -----------------------------
# Helpers: Load data
# -----------------------------
def _clean_phone(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\D+", "", str(s))


def load_students_df() -> pd.DataFrame:
    """
    يتوقع وجود أعمدة (يفضّل):
      training_id, name, major, phone
    ممكن تكون بالعربي أيضًا:
      الرقم_التدريبي, الاسم, التخصص, الجوال
    """
    df = None
    if STUDENTS_XLSX.exists():
        df = pd.read_excel(STUDENTS_XLSX)
    elif STUDENTS_CSV.exists():
        df = pd.read_csv(STUDENTS_CSV)
    else:
        # إذا ما عندك ملف طلاب، رجّع df فاضي مع أعمدة افتراضية
        return pd.DataFrame(columns=["training_id", "name", "major", "phone"])

    # توحيد أسماء الأعمدة
    cols_map = {}
    for c in df.columns:
        c2 = str(c).strip()
        if c2 in ["training_id", "الرقم_التدريبي", "رقم_تدريبي", "رقم التدريب", "رقم_التدريب"]:
            cols_map[c] = "training_id"
        elif c2 in ["name", "الاسم", "اسم", "اسم_الطالب"]:
            cols_map[c] = "name"
        elif c2 in ["major", "التخصص", "تخصص", "القسم"]:
            cols_map[c] = "major"
        elif c2 in ["phone", "الجوال", "رقم_الجوال", "الهاتف"]:
            cols_map[c] = "phone"

    df = df.rename(columns=cols_map)

    # تأكد من الأعمدة
    for needed in ["training_id", "name", "major", "phone"]:
        if needed not in df.columns:
            df[needed] = ""

    df["training_id"] = df["training_id"].astype(str).str.strip()
    df["phone"] = df["phone"].apply(_clean_phone)
    return df


def load_entities() -> list[str]:
    """
    يتوقع وجود عمود:
      entity  أو  الجهة
    """
    if ENTITIES_XLSX.exists():
        df = pd.read_excel(ENTITIES_XLSX)
    elif ENTITIES_CSV.exists():
        df = pd.read_csv(ENTITIES_CSV)
    else:
        # قائمة افتراضية إذا ما عندك ملف جهات
        return [
            "مدير مركز النصر لصيانة السيارات (المتبقي: 5)",
            "شركة مثال للتقنية (المتبقي: 5)",
            "مؤسسة مثال للخدمات (المتبقي: 5)",
        ]

    # تحديد العمود
    col = None
    for c in df.columns:
        if str(c).strip() in ["entity", "الجهة", "جهة_التدريب", "جهة التدريب"]:
            col = c
            break
    if col is None:
        # إذا ما لقينا عمود، خذ أول عمود
        col = df.columns[0]

    entities = [str(x).strip() for x in df[col].dropna().tolist() if str(x).strip()]
    return entities


def find_student(training_id: str, last4: str) -> dict | None:
    df = load_students_df()
    training_id = str(training_id).strip()
    last4 = re.sub(r"\D+", "", str(last4).strip())

    if df.empty:
        return None

    row = df[df["training_id"] == training_id]
    if row.empty:
        return None

    phone = row.iloc[0]["phone"] or ""
    if len(phone) < 4:
        return None

    if phone[-4:] != last4:
        return None

    return {
        "training_id": training_id,
        "name": str(row.iloc[0]["name"]).strip(),
        "major": str(row.iloc[0]["major"]).strip(),
        "phone": phone,
    }


# -----------------------------
# Helpers: DOCX -> PDF
# -----------------------------
def _find_soffice_cmd() -> str | None:
    # أشهر أماكن/أسماء
    candidates = [
        "soffice",
        "libreoffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
    ]
    for c in candidates:
        if shutil.which(c) or Path(c).exists():
            return c
    return None


def convert_docx_to_pdf(docx_path: Path, pdf_path: Path) -> None:
    """
    يعتمد على LibreOffice داخل Docker (اللي ثبتته في Dockerfile)
    """
    soffice = _find_soffice_cmd()
    if not soffice:
        raise FileNotFoundError("LibreOffice (soffice) غير موجود داخل السيرفر.")

    outdir = pdf_path.parent
    outdir.mkdir(parents=True, exist_ok=True)

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(outdir),
        str(docx_path),
    ]

    res = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"فشل التحويل إلى PDF:\n{res.stderr}\n{res.stdout}")

    # LibreOffice ينتج PDF باسم docx نفسه
    produced = outdir / (docx_path.stem + ".pdf")
    if not produced.exists():
        raise FileNotFoundError("لم يتم إنشاء ملف PDF بعد التحويل.")
    if produced != pdf_path:
        produced.replace(pdf_path)


def build_letter_docx(student: dict, entity_name: str, out_docx: Path) -> None:
    """
    يصنع DOCX بسيط بخطاب توجيه.
    """
    doc = Document()

    # عنوان
    p = doc.add_paragraph("خطاب توجيه تدريب تعاوني")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.runs[0]
    run.bold = True
    run.font.size = Pt(18)

    doc.add_paragraph("")

    # متن الخطاب
    body = [
        f"التاريخ: {datetime.now().strftime('%Y/%m/%d')}",
        f"سعادة/ {entity_name}",
        "",
        "السلام عليكم ورحمة الله وبركاته،",
        "",
        "نفيدكم بأن الطالب الموضحة بياناته أدناه أحد طلاب الكلية التقنية، ويرغب في تنفيذ التدريب التعاوني لديكم.",
        "",
        f"الاسم: {student.get('name','')}",
        f"الرقم التدريبي: {student.get('training_id','')}",
        f"التخصص: {student.get('major','')}",
        f"الجوال: {student.get('phone','')}",
        "",
        "نأمل منكم التكرم بالتعاون وتسهيل مهمة التدريب للطالب، شاكرين لكم تعاونكم.",
        "",
        "وتقبلوا خالص التحية والتقدير.",
    ]

    for line in body:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT  # يمين (RTL بصرياً)

    doc.save(out_docx)


# -----------------------------
# HTML Pages
# -----------------------------
PAGE_LOGIN = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
    .top-image img{width:100%;height:100%;object-fit:cover;display:block}
    .wrap{max-width:900px;margin:30px auto;padding:0 16px}
    .card{background:#fff;border-radius:18px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:26px}
    h1{margin:0 0 10px 0;font-size:34px}
    p{color:#444;margin:0 0 18px 0}
    .row{display:flex;gap:12px;flex-wrap:wrap}
    .field{flex:1;min-width:220px}
    label{display:block;margin:10px 0 6px 0;font-weight:700}
    input{width:100%;padding:12px;border:1px solid #ddd;border-radius:12px;font-size:16px}
    .btn{margin-top:16px;width:100%;padding:14px;border:none;border-radius:12px;background:#0f172a;color:#fff;font-size:18px;cursor:pointer}
    .err{color:#b91c1c;margin-top:12px}
    .hint{color:#777;font-size:14px;margin-top:10px}
  </style>
</head>
<body>

  <div class="top-image">
    <img src="/static/{{header_img}}" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

      <form method="POST" action="/login">
        <div class="row">
          <div class="field">
            <label>الرقم التدريبي</label>
            <input name="training_id" required placeholder="مثال: 444242291" value="{{training_id or ''}}">
          </div>
          <div class="field">
            <label>آخر 4 أرقام من الجوال</label>
            <input name="last4" required placeholder="مثال: 5513" maxlength="4" value="{{last4 or ''}}">
          </div>
        </div>

        <button class="btn" type="submit">دخول</button>

        {% if error %}
          <div class="err">{{error}}</div>
        {% endif %}

        <div class="hint">
          ملاحظة: بيانات التحقق تُقرأ من ملف الطلاب داخل مجلد <b>data</b>.
        </div>
      </form>
    </div>
  </div>

</body>
</html>
"""


PAGE_HOME = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
    .top-image img{width:100%;height:100%;object-fit:cover;display:block}
    .wrap{max-width:1000px;margin:30px auto;padding:0 16px}
    .card{background:#fff;border-radius:18px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:26px}
    h1{margin:0 0 10px 0;font-size:38px;text-align:center}
    .sub{margin:0 0 22px 0;text-align:center;color:#333}
    .info{background:#eef6ff;border-radius:14px;padding:14px 16px;margin:18px 0;color:#0b3a77}
    .info div{margin:6px 0;font-weight:700}
    label{display:block;margin:16px 0 8px 0;font-weight:700}
    select{width:100%;padding:12px;border:1px solid #ddd;border-radius:12px;font-size:16px}
    .btn{margin-top:18px;width:100%;padding:14px;border:none;border-radius:12px;background:#0f172a;color:#fff;font-size:18px;cursor:pointer}
    .links{display:flex;justify-content:space-between;margin-top:16px}
    .links a{color:#0b3a77;text-decoration:none;font-weight:700}
    .note{color:#777;font-size:14px;margin-top:14px;text-align:center}
    .err{color:#b91c1c;margin-top:10px;text-align:center;font-weight:700}
  </style>
</head>
<body>

  <div class="top-image">
    <img src="/static/{{header_img}}" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>{{title}}</h1>
      <p class="sub">مرحباً بك في نظام إصدار خطاب التدريب التعاوني</p>

      <div class="info">
        <div>الاسم: {{student.name}}</div>
        <div>الرقم التدريبي: {{student.training_id}}</div>
        <div>التخصص: {{student.major}}</div>
        <div>الجوال: {{student.phone}}</div>
      </div>

      <form method="POST" action="/generate">
        <label>اختر جهة التدريب المتاحة:</label>
        <select name="entity" required>
          <option value="">-- اختر --</option>
          {% for e in entities %}
            <option value="{{e}}">{{e}}</option>
          {% endfor %}
        </select>

        <button class="btn" type="submit">طباعة خطاب التوجيه PDF</button>
      </form>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}

      <div class="note">
        إذا ظهر خطأ عند الطباعة: افتح Logs في Render ثم Live tail وأرسل آخر سطرين من الخطأ.
      </div>

      <div class="links">
        <a href="/logout">تسجيل خروج</a>
        <a href="/">رجوع</a>
      </div>
    </div>
  </div>

</body>
</html>
"""


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def index():
    # لو مسجل دخول -> روح للصفحة الرئيسية
    if session.get("student"):
        return redirect(url_for("home"))
    return redirect(url_for("login"))


@app.get("/login")
def login():
    return render_template_string(
        PAGE_LOGIN,
        title=APP_TITLE,
        header_img=HEADER_IMG,
        error=None,
        training_id="",
        last4="",
    )


@app.post("/login")
def do_login():
    training_id = request.form.get("training_id", "").strip()
    last4 = request.form.get("last4", "").strip()

    student = find_student(training_id, last4)
    if not student:
        return render_template_string(
            PAGE_LOGIN,
            title=APP_TITLE,
            header_img=HEADER_IMG,
            error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.",
            training_id=training_id,
            last4=last4,
        )

    session["student"] = student
    return redirect(url_for("home"))


@app.get("/home")
def home():
    student = session.get("student")
    if not student:
        return redirect(url_for("login"))

    entities = load_entities()[:50]  # تقدر تقللها
    return render_template_string(
        PAGE_HOME,
        title=APP_TITLE,
        header_img=HEADER_IMG,
        student=student,
        entities=entities,
        error=None,
    )


@app.post("/generate")
def generate():
    student = session.get("student")
    if not student:
        return redirect(url_for("login"))

    entity = request.form.get("entity", "").strip()
    if not entity:
        entities = load_entities()[:50]
        return render_template_string(
            PAGE_HOME,
            title=APP_TITLE,
            header_img=HEADER_IMG,
            student=student,
            entities=entities,
            error="اختر جهة التدريب أولاً.",
        )

    # ملفات الإخراج
    token = uuid.uuid4().hex[:10]
    out_docx = OUT_DIR / f"letter_{student['training_id']}_{token}.docx"
    out_pdf = OUT_DIR / f"letter_{student['training_id']}_{token}.pdf"

    try:
        build_letter_docx(student, entity, out_docx)
        convert_docx_to_pdf(out_docx, out_pdf)
    except Exception as e:
        print("ERROR /generate:", repr(e))
        entities = load_entities()[:50]
        return render_template_string(
            PAGE_HOME,
            title=APP_TITLE,
            header_img=HEADER_IMG,
            student=student,
            entities=entities,
            error=f"حدث خطأ أثناء توليد PDF: {e}",
        )

    # تنزيل PDF
    return send_file(
        out_pdf,
        as_attachment=True,
        download_name=f"خطاب_توجيه_{student['training_id']}.pdf",
        mimetype="application/pdf",
    )


@app.get("/logout")
def logout():
    session.pop("student", None)
    return redirect(url_for("login"))


# Render / Docker: التشغيل على PORT
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port, debug=False)

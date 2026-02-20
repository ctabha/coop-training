import os
import json
import uuid
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, render_template_string

# =========================
# إعدادات عامة
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"
OUT_DIR.mkdir(exist_ok=True)

STUDENTS_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"  # سيتم إنشاؤه تلقائياً إذا غير موجود

# =========================
# أدوات مساعدة
# =========================
AR_DIACRITICS = str.maketrans("", "", "ًٌٍَُِّْـ")  # تشكيل + تطويل


def norm_col(s: str) -> str:
    """توحيد اسم العمود: إزالة مسافات/تشكيل/تطويل"""
    if s is None:
        return ""
    s = str(s).strip()
    s = s.translate(AR_DIACRITICS)
    s = s.replace(" ", "").replace("\u00A0", "")
    return s


def load_json(path: Path, default):
    if path.exists():
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return default
    return default


def save_json(path: Path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


def load_students() -> pd.DataFrame:
    if not STUDENTS_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {STUDENTS_FILE}")

    df = pd.read_excel(STUDENTS_FILE)

    # خريطة الأعمدة المطلوبة (نعرفها عبر norm_col لتجاوز اختلافات المسافات مثل "جهة التدريب ")
    col_map = {}
    for c in df.columns:
        col_map[norm_col(c)] = c

    def pick(*candidates):
        for k in candidates:
            nk = norm_col(k)
            if nk in col_map:
                return col_map[nk]
        return None

    required = {
        "trainee_id": pick("رقم المتدرب", "رقمالمتدرب"),
        "trainee_name": pick("إسم المتدرب", "اسم المتدرب", "اسمالمتدرب"),
        "phone": pick("رقم الجوال", "رقمالجوال", "الجوال"),
        "specialization": pick("التخصص"),
        "program": pick("برنامج"),
        "entity": pick("جهة التدريب", "جهةالتدريب"),
        "department": pick("القسم"),
        "trainer": pick("المدرب"),
        "course_name": pick("اسم المقرر"),
        "ref": pick("الرقم المرجعي"),
    }

    missing = [k for k, v in required.items() if v is None]
    if missing:
        raise ValueError(
            "لم أستطع العثور على أعمدة مطلوبة في ملف الإكسل. "
            f"المفقودة: {missing}\n"
            f"الأعمدة الموجودة: {list(df.columns)}"
        )

    # إعادة تسمية لأسماء داخلية ثابتة
    df = df.rename(columns={
        required["trainee_id"]: "trainee_id",
        required["trainee_name"]: "trainee_name",
        required["phone"]: "phone",
        required["specialization"]: "specialization",
        required["program"]: "program",
        required["entity"]: "entity",
        required["department"]: "department",
        required["trainer"]: "trainer",
        required["course_name"]: "course_name",
        required["ref"]: "ref",
    })

    # تنظيف
    df["entity"] = df["entity"].astype(str).str.strip()
    df["specialization"] = df["specialization"].astype(str).str.strip()
    df["program"] = df["program"].astype(str).str.strip()
    df["trainee_name"] = df["trainee_name"].astype(str).str.strip()

    # تأكد من أن رقم الجوال نص
    df["phone"] = df["phone"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df["trainee_id"] = df["trainee_id"].astype(str).str.replace(".0", "", regex=False).str.strip()

    return df


def last4(s: str) -> str:
    s = "".join([ch for ch in str(s) if ch.isdigit()])
    return s[-4:] if len(s) >= 4 else s


def compute_capacity(df: pd.DataFrame) -> dict:
    """
    الفرص = عدد تكرار (التخصص + جهة التدريب) في ملف الاكسل
    يرجع dict بالشكل:
    capacity[specialization][entity] = count
    """
    cap = {}
    grp = df.groupby(["specialization", "entity"]).size().reset_index(name="count")
    for _, row in grp.iterrows():
        sp = row["specialization"]
        ent = row["entity"]
        cnt = int(row["count"])
        cap.setdefault(sp, {})[ent] = cnt
    return cap


def compute_remaining(capacity: dict, assignments: dict) -> dict:
    """
    remaining = capacity - chosen
    assignments format:
    {
      "trainee_id": {"entity": "...", "specialization": "...", "ts": "..."}
    }
    """
    remaining = {sp: ent_map.copy() for sp, ent_map in capacity.items()}

    # اطرح الاختيارات
    for tid, info in assignments.items():
        sp = info.get("specialization")
        ent = info.get("entity")
        if sp in remaining and ent in remaining[sp]:
            remaining[sp][ent] = max(0, int(remaining[sp][ent]) - 1)

    # حذف الجهات التي أصبحت 0
    cleaned = {}
    for sp, ent_map in remaining.items():
        cleaned[sp] = {e: n for e, n in ent_map.items() if int(n) > 0}
    return cleaned


# =========================
# تطبيق Flask
# =========================
app = Flask(__name__)


PAGE_LOGIN = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{
      font-family: Arial, sans-serif;
      background:#f7f7f7;
      margin:0;
    }
    .top-image{
      width:100%;
      height:25vh;            /* ربع الصفحة تقريباً */
      background:#fff;
      display:flex;
      align-items:center;
      justify-content:center;
      border-bottom:1px solid #eee;
    }
    .top-image img{
      width:100%;
      height:100%;
      object-fit:contain;     /* مهم: يمنع قص الصورة */
    }
    .wrap{
      display:flex;
      align-items:center;
      justify-content:center;
      padding:30px 12px 60px;
    }
    .card{
      width:min(900px, 95vw);
      background:#fff;
      border-radius:20px;
      box-shadow:0 10px 25px rgba(0,0,0,0.08);
      padding:30px;
      text-align:center;
    }
    h1{ margin: 10px 0 10px; font-size:42px; }
    p{ margin: 0 0 20px; color:#444; font-size:18px; }
    .row{
      display:flex;
      gap:16px;
      margin:20px 0;
      flex-wrap:wrap;
    }
    .field{
      flex:1;
      min-width:260px;
      text-align:right;
    }
    label{ display:block; margin-bottom:8px; font-weight:bold; }
    input{
      width:100%;
      padding:14px 14px;
      border:1px solid #ddd;
      border-radius:14px;
      font-size:18px;
      outline:none;
      background:#fff;
    }
    .btn{
      width:100%;
      padding:18px;
      border:none;
      border-radius:18px;
      background:#0b1730;
      color:#fff;
      font-size:22px;
      cursor:pointer;
      margin-top:10px;
    }
    .err{
      color:#c00;
      margin-top:14px;
      white-space:pre-wrap;
      text-align:center;
      font-size:16px;
    }
    .note{
      color:#666;
      margin-top:8px;
      font-size:14px;
    }
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
            <input name="last4" placeholder="مثال: 6101" required>
          </div>
        </div>

        <button class="btn" type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="err">{{ error }}</div>
      {% endif %}

      <div class="note">ملاحظة: يتم التحقق من ملف الطلاب داخل مجلد data.</div>
    </div>
  </div>
</body>
</html>
"""


PAGE_DASH = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial; background:#f7f7f7; margin:0;}
    .top-image{
      width:100%; height:25vh; background:#fff;
      display:flex; align-items:center; justify-content:center;
      border-bottom:1px solid #eee;
    }
    .top-image img{width:100%; height:100%; object-fit:contain;}
    .wrap{display:flex; justify-content:center; padding:25px 12px 60px;}
    .card{
      width:min(1000px, 95vw);
      background:#fff; border-radius:20px;
      box-shadow:0 10px 25px rgba(0,0,0,.08);
      padding:28px;
    }
    h2{margin:0 0 10px;}
    .muted{color:#555;}
    .box{
      background:#f3f5f8; border-radius:16px; padding:16px; margin-top:16px;
      text-align:right;
    }
    select{
      width:100%; padding:14px; border-radius:14px; border:1px solid #ddd;
      font-size:16px; background:#fff;
    }
    .btn{
      width:100%; padding:16px; border:none; border-radius:16px;
      background:#0b1730; color:#fff; font-size:18px; cursor:pointer;
      margin-top:10px;
    }
    .ok{color:green; margin-top:10px;}
    .err{color:#c00; margin-top:10px; white-space:pre-wrap;}
    .grid{display:grid; grid-template-columns:1fr 1fr; gap:12px;}
    @media (max-width:900px){ .grid{grid-template-columns:1fr;} }
    ul{margin:10px 0 0; padding-right:18px;}
    a{color:#0b1730; font-weight:bold;}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h2>مرحباً {{ trainee_name }}</h2>
      <div class="muted">
        التخصص: <b>{{ specialization }}</b> — البرنامج: <b>{{ program }}</b>
      </div>

      {% if already %}
        <div class="box">
          <b>تم تسجيل اختيارك مسبقاً:</b><br>
          الجهة المختارة: <b>{{ chosen_entity }}</b><br><br>
          <a href="/letter?tid={{ trainee_id }}">تحميل/طباعة خطاب التوجيه PDF</a>
        </div>
      {% else %}
        <div class="box">
          <b>اختر جهة التدريب المتاحة لتخصصك (تختفي الجهة إذا انتهت فرصها):</b>
          <form method="post" action="/choose">
            <input type="hidden" name="tid" value="{{ trainee_id }}">
            <select name="entity" required>
              <option value="" disabled selected>-- اختر الجهة --</option>
              {% for ent, rem in options %}
                <option value="{{ ent }}">{{ ent }} (متبقي: {{ rem }})</option>
              {% endfor %}
            </select>
            <button class="btn" type="submit">تأكيد الاختيار</button>
          </form>

          {% if msg_ok %}<div class="ok">{{ msg_ok }}</div>{% endif %}
          {% if error %}<div class="err">{{ error }}</div>{% endif %}
        </div>
      {% endif %}

      <div class="box">
        <b>ملخص الفرص المتبقية حسب الجهات داخل تخصصك:</b>
        <ul>
          {% for ent, rem in options %}
            <li>{{ ent }}: {{ rem }} فرصة</li>
          {% endfor %}
        </ul>
      </div>

      <div class="box">
        <b>ملخص الفرص المتبقية حسب التخصصات (للإدارة):</b>
        <div class="grid">
          {% for sp, total in spec_totals %}
            <div>• {{ sp }}: <b>{{ total }}</b> فرصة متبقية</div>
          {% endfor %}
        </div>
      </div>

    </div>
  </div>
</body>
</html>
"""


def build_letter_pdf(student_row: dict, chosen_entity: str) -> Path:
    """
    إنشاء DOCX بسيط ثم تحويله إلى PDF.
    (إذا عندك قالب letter_template.docx وتبغى دمج حقول، قول لي وأعدله لك)
    """
    # نص عربي بسيط داخل DOCX عبر python-docx
    from docx import Document

    doc = Document()
    doc.add_paragraph("خطاب توجيه التدريب التعاوني")
    doc.add_paragraph("")
    doc.add_paragraph(f"اسم المتدرب: {student_row.get('trainee_name','')}")
    doc.add_paragraph(f"رقم المتدرب: {student_row.get('trainee_id','')}")
    doc.add_paragraph(f"التخصص: {student_row.get('specialization','')}")
    doc.add_paragraph(f"البرنامج: {student_row.get('program','')}")
    doc.add_paragraph(f"الجهة التدريبية: {chosen_entity}")
    doc.add_paragraph("")
    doc.add_paragraph("مع تمنياتنا لكم بالتوفيق.")

    token = uuid.uuid4().hex
    docx_path = OUT_DIR / f"letter_{token}.docx"
    pdf_path = OUT_DIR / f"letter_{token}.pdf"
    doc.save(docx_path)

    # تحويل إلى PDF بواسطة LibreOffice
    # يحتاج وجود libreoffice داخل Docker (عندك مثبت سابقاً)
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(OUT_DIR), str(docx_path)],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    if not pdf_path.exists():
        # LibreOffice أحياناً يسمي الملف بنفس اسم docx
        alt_pdf = OUT_DIR / (docx_path.stem + ".pdf")
        if alt_pdf.exists():
            alt_pdf.rename(pdf_path)

    return pdf_path


# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    error = None

    if request.method == "POST":
        tid = (request.form.get("trainee_id") or "").strip()
        l4 = (request.form.get("last4") or "").strip()

        try:
            df = load_students()
        except Exception as e:
            return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=str(e))

        # تحقق
        row = df[df["trainee_id"] == tid]
        if row.empty:
            error = "الرقم التدريبي غير موجود."
        else:
            phone = row.iloc[0]["phone"]
            if last4(phone) != last4(l4):
                error = "آخر 4 أرقام من الجوال غير صحيحة."
            else:
                return redirect(url_for("dashboard", tid=tid))

    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error)


@app.route("/dashboard")
def dashboard():
    tid = (request.args.get("tid") or "").strip()

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("index"))

    student = row.iloc[0].to_dict()
    specialization = student["specialization"]

    capacity = compute_capacity(df)
    assignments = load_json(ASSIGNMENTS_FILE, default={})
    remaining = compute_remaining(capacity, assignments)

    # خيارات هذا التخصص فقط
    options_map = remaining.get(specialization, {})
    options = sorted(options_map.items(), key=lambda x: (-int(x[1]), x[0]))

    # إجمالي متبقي لكل تخصص
    spec_totals = []
    for sp, ent_map in remaining.items():
        spec_totals.append((sp, int(sum(ent_map.values()))))
    spec_totals.sort(key=lambda x: -x[1])

    already = tid in assignments
    chosen_entity = assignments.get(tid, {}).get("entity") if already else None

    return render_template_string(
        PAGE_DASH,
        title=APP_TITLE,
        trainee_id=tid,
        trainee_name=student.get("trainee_name", ""),
        specialization=specialization,
        program=student.get("program", ""),
        options=options,
        spec_totals=spec_totals,
        already=already,
        chosen_entity=chosen_entity,
        msg_ok=None,
        error=None
    )


@app.route("/choose", methods=["POST"])
def choose():
    tid = (request.form.get("tid") or "").strip()
    entity = (request.form.get("entity") or "").strip()

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("index"))

    student = row.iloc[0].to_dict()
    specialization = student["specialization"]

    capacity = compute_capacity(df)
    assignments = load_json(ASSIGNMENTS_FILE, default={})
    remaining = compute_remaining(capacity, assignments)

    # إذا سبق اختار
    if tid in assignments:
        return redirect(url_for("dashboard", tid=tid))

    # تحقق أن الجهة متاحة لهذا التخصص وما زال فيها فرص
    if entity not in remaining.get(specialization, {}):
        # انتهت أو ليست ضمن تخصصه
        return render_template_string(
            PAGE_DASH,
            title=APP_TITLE,
            trainee_id=tid,
            trainee_name=student.get("trainee_name", ""),
            specialization=specialization,
            program=student.get("program", ""),
            options=sorted(remaining.get(specialization, {}).items(), key=lambda x: (-int(x[1]), x[0])),
            spec_totals=sorted([(sp, int(sum(ent_map.values()))) for sp, ent_map in remaining.items()], key=lambda x: -x[1]),
            already=False,
            chosen_entity=None,
            msg_ok=None,
            error="هذه الجهة غير متاحة الآن (قد تكون انتهت فرصها أو ليست ضمن تخصصك)."
        )

    # احفظ الاختيار
    assignments[tid] = {
        "entity": entity,
        "specialization": specialization,
        "ts": datetime.now().isoformat(timespec="seconds")
    }
    save_json(ASSIGNMENTS_FILE, assignments)

    return redirect(url_for("dashboard", tid=tid))


@app.route("/letter")
def letter():
    tid = (request.args.get("tid") or "").strip()
    assignments = load_json(ASSIGNMENTS_FILE, default={})
    if tid not in assignments:
        return redirect(url_for("index"))

    chosen_entity = assignments[tid]["entity"]

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("index"))

    student = row.iloc[0].to_dict()

    pdf_path = build_letter_pdf(student, chosen_entity)
    return send_file(pdf_path, as_attachment=True, download_name="خطاب_التوجيه.pdf")


@app.route("/health")
def health():
    return {"ok": True}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

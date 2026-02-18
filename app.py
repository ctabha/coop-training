import os
import json
import uuid
import traceback
import subprocess
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, send_file

from docx import Document  # python-docx

app = Flask(__name__)

# =========================
# Paths & Config
# =========================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

# ملفات البيانات
STUDENTS_XLSX = DATA_DIR / "students.xlsx"  # اسم ملفك الصحيح
TEMPLATE_DOCX = DATA_DIR / "letter_template.docx"  # قالب الخطاب
SLOTS_FILE = DATA_DIR / "slots_by_specialty.json"  # تخزين الفرص المتبقية

# =========================
# Helpers
# =========================
def last4_digits(x) -> str:
    s = "".join(ch for ch in str(x) if ch.isdigit())
    return s[-4:] if len(s) >= 4 else s

def norm_key(s: str) -> str:
    s = str(s).strip().lower()
    for ch in [" ", "\t", "\n", "\r", "-", "_", ".", "ـ", "(", ")", "[", "]", "{", "}", ":", "؛", ",", "،", "/", "\\"]:
        s = s.replace(ch, "")
    return s

def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    cols = [str(c) for c in df.columns]
    norm_map = {norm_key(c): c for c in cols}

    # تطابق مباشر
    for cand in candidates:
        k = norm_key(cand)
        if k in norm_map:
            return norm_map[k]

    # تطابق جزئي
    for cand in candidates:
        kc = norm_key(cand)
        for real in cols:
            if kc and kc in norm_key(real):
                return real

    raise KeyError(f"لم أجد عمود من: {candidates}\nالأعمدة الموجودة: {cols}")

def load_students_df() -> pd.DataFrame:
    """
    يقرأ students.xlsx ويختار الشيت الصحيح تلقائياً:
    يبحث عن شيت يحتوي (رقم المتدرب + رقم الجوال + اسم المتدرب).
    """
    if not STUDENTS_XLSX.exists():
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_XLSX}")

    xls = pd.ExcelFile(STUDENTS_XLSX)
    best_df = None
    best_score = -1
    best_sheet = None

    # مرشحات أسماء الأعمدة
    id_candidates = ["رقم المتدرب", "رقم_المتدرب", "StudentID", "ID"]
    mobile_candidates = ["رقم الجوال", "الجوال", "Mobile", "Phone"]
    name_candidates = ["اسم المتدرب", "اسم_المتدرب", "الاسم", "Name"]
    spec_candidates = ["التخصص", "تخصص", "برنامج", "البرنامج", "Specialty", "Major", "Program"]
    entity_candidates = ["جهة التدريب", "الجهة", "جهة", "TrainingEntity", "Entity", "Company"]

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(STUDENTS_XLSX, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]
            cols_norm = {norm_key(c) for c in df.columns}

            score = 0
            # كل واحد موجود يزيد نقاط
            def has_any(cands):
                for cc in cands:
                    if norm_key(cc) in cols_norm:
                        return True
                    # فحص جزئي
                    for real in df.columns:
                        if norm_key(cc) in norm_key(real):
                            return True
                return False

            if has_any(id_candidates): score += 3
            if has_any(mobile_candidates): score += 3
            if has_any(name_candidates): score += 3
            if has_any(spec_candidates): score += 1
            if has_any(entity_candidates): score += 1

            if score > best_score:
                best_score = score
                best_df = df
                best_sheet = sheet

        except Exception:
            continue

    if best_df is None:
        raise RuntimeError("لم أستطع قراءة أي Sheet من ملف الطلاب.")

    # إذا كان أفضل شيت نقاطه ضعيفة، غالباً الشيت غلط
    if best_score < 7:
        raise RuntimeError(
            f"تم العثور على Sheet لكن لا يحتوي الأعمدة المطلوبة.\n"
            f"أفضل Sheet: {best_sheet} | Score={best_score}\n"
            f"الأعمدة: {list(best_df.columns)}"
        )

    return best_df

# =========================
# Slots logic (فرص حسب التخصص)
# =========================
def calculate_slots_from_excel() -> dict:
    """
    يحسب الفرص المتاحة لكل تخصص من ملف students.xlsx:
    - يحدد التخصص (التخصص أو برنامج)
    - يحدد جهة التدريب
    - الفرص = عدد تكرار كل جهة داخل نفس التخصص
    """
    df = load_students_df()

    col_id = find_col(df, ["رقم المتدرب", "StudentID", "ID"])
    col_mobile = find_col(df, ["رقم الجوال", "الجوال", "Mobile", "Phone"])
    col_name = find_col(df, ["اسم المتدرب", "الاسم", "Name"])

    # تخصص/برنامج
    try:
        col_spec = find_col(df, ["التخصص", "تخصص", "برنامج", "البرنامج", "Specialty", "Major", "Program"])
    except Exception:
        col_spec = "__ALL__"
        df[col_spec] = "عام"

    # جهة التدريب
    try:
        col_entity = find_col(df, ["جهة التدريب", "الجهة", "جهة", "TrainingEntity", "Entity", "Company"])
    except Exception:
        # إذا لم توجد جهة تدريب، لا يمكن حساب فرص
        raise KeyError("لا يوجد عمود 'جهة التدريب' في ملف الطلاب. لا يمكن حساب الفرص.")

    df[col_spec] = df[col_spec].astype(str).str.strip()
    df[col_entity] = df[col_entity].astype(str).str.strip()

    # إزالة الصفوف الفارغة
    df = df[(df[col_entity].notna()) & (df[col_entity] != "") & (df[col_spec].notna()) & (df[col_spec] != "")]

    slots = {}
    for spec, g in df.groupby(col_spec):
        counts = g[col_entity].value_counts().to_dict()
        slots[str(spec)] = {str(ent): int(cnt) for ent, cnt in counts.items()}

    return slots

def load_slots_by_specialty() -> dict:
    """
    يقرأ slots_by_specialty.json إن وجد، وإلا يحسبه من excel.
    """
    if SLOTS_FILE.exists():
        with open(SLOTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)

    slots = calculate_slots_from_excel()
    save_slots_by_specialty(slots)
    return slots

def save_slots_by_specialty(slots: dict) -> None:
    with open(SLOTS_FILE, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

def specialty_total_remaining(slots: dict, specialty: str) -> int:
    spec_slots = slots.get(specialty, {})
    return sum(int(v) for v in spec_slots.values() if int(v) > 0)

# =========================
# DOCX placeholders -> PDF
# =========================
def _replace_in_paragraph(paragraph, mapping: dict[str, str]) -> None:
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return
    replaced = full_text
    for k, v in mapping.items():
        replaced = replaced.replace(k, v)
    if replaced == full_text:
        return
    for run in paragraph.runs:
        run.text = ""
    if paragraph.runs:
        paragraph.runs[0].text = replaced
    else:
        paragraph.add_run(replaced)

def _replace_in_table(table, mapping: dict[str, str]) -> None:
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                _replace_in_paragraph(p, mapping)

def render_docx_to_pdf(template_path: Path, out_pdf_path: Path, mapping: dict[str, str]) -> None:
    if not template_path.exists():
        raise FileNotFoundError(f"قالب الخطاب غير موجود: {template_path}")

    tmp_docx = OUT_DIR / f"tmp_{uuid.uuid4().hex}.docx"
    doc = Document(str(template_path))

    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)
    for t in doc.tables:
        _replace_in_table(t, mapping)

    doc.save(str(tmp_docx))

    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(OUT_DIR),
        str(tmp_docx),
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    if result.returncode != 0:
        raise RuntimeError(
            "فشل تحويل DOCX إلى PDF عبر LibreOffice.\n"
            f"STDOUT:\n{result.stdout}\n\nSTDERR:\n{result.stderr}"
        )

    produced_pdf = OUT_DIR / (tmp_docx.stem + ".pdf")
    if not produced_pdf.exists():
        raise FileNotFoundError("تم تشغيل التحويل لكن ملف PDF لم يُنتج.")

    produced_pdf.replace(out_pdf_path)

    try:
        if tmp_docx.exists():
            tmp_docx.unlink()
    except Exception:
        pass

# =========================
# HTML Templates (مع هيدر بدون قص)
# =========================
LOGIN_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f4f4f4;margin:0}
.top-image{
  width:100%;
  height:25vh;            /* ربع الشاشة */
  background:#fff;
  display:flex;
  align-items:center;
  justify-content:center;
  overflow:hidden;
}
.top-image img{
  max-height:100%;
  max-width:100%;
  width:auto;
  height:auto;
  object-fit:contain;     /* يمنع القص */
  display:block;
}
.container{
  max-width:900px;
  margin:-30px auto 50px;
  background:#fff;
  padding:35px;
  border-radius:16px;
  text-align:center;
  box-shadow:0 10px 25px rgba(0,0,0,.08)
}
.row{display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin-top:18px}
.field{flex:1;min-width:260px;text-align:right}
label{display:block;font-weight:700;margin:10px 0}
input{width:100%;padding:14px;border-radius:12px;border:1px solid #ddd;font-size:16px}
button{width:100%;background:#0b1a3a;color:#fff;padding:16px 24px;border:none;border-radius:14px;font-size:18px;margin-top:18px;cursor:pointer}
.error{color:#c00;margin-top:14px;font-weight:700;white-space:pre-wrap;text-align:right}
.small{color:#666;margin-top:10px}
.links{margin-top:10px;color:#666;font-size:14px}
.links a{color:#0b1a3a;text-decoration:none}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h1 style="margin:0">{{title}}</h1>
  <p class="small">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

  <form method="post">
    <div class="row">
      <div class="field">
        <label>الرقم التدريبي</label>
        <input name="training_number" placeholder="مثال: 444229747" required>
      </div>
      <div class="field">
        <label>آخر 4 أرقام من الجوال</label>
        <input name="mobile_last4" placeholder="مثال: 6101" required>
      </div>
    </div>
    <button type="submit">دخول</button>
  </form>

  <div class="links">
    <a href="/slots">عرض الفرص المتاحة حسب التخصص</a>
    •
    <a href="/recalc">إعادة حساب الفرص من ملف الإكسل</a>
  </div>

  {% if error %}
    <div class="error">{{error}}</div>
  {% endif %}
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
<style>
body{font-family:Arial;background:#f4f4f4;margin:0}
.top-image{
  width:100%;
  height:25vh;
  background:#fff;
  display:flex;
  align-items:center;
  justify-content:center;
  overflow:hidden;
}
.top-image img{
  max-height:100%;
  max-width:100%;
  width:auto;
  height:auto;
  object-fit:contain;
  display:block;
}
.container{
  max-width:900px;
  margin:-30px auto 50px;
  background:#fff;
  padding:35px;
  border-radius:16px;
  text-align:center;
  box-shadow:0 10px 25px rgba(0,0,0,.08)
}
select{width:100%;padding:14px;border-radius:12px;border:1px solid #ddd;font-size:16px;margin-top:10px}
button{width:100%;background:#0b1a3a;color:#fff;padding:16px 24px;border:none;border-radius:14px;font-size:18px;margin-top:18px;cursor:pointer}
.note{color:#666;margin-top:10px}
.warn{color:#c00;font-weight:700;margin-top:14px;white-space:pre-wrap;text-align:right}
.badge{display:inline-block;background:#eef2ff;color:#0b1a3a;padding:6px 12px;border-radius:999px;font-weight:700;margin-top:10px}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h2 style="margin:0">مرحباً {{name}}</h2>
  <div class="badge">تخصصك: {{specialty}} • المتبقي: {{total_remaining}} فرصة</div>

  {% if no_entities %}
    <div class="warn">لا توجد جهات تدريبية متاحة لهذا التخصص حالياً.</div>
  {% else %}
    <form method="post" action="/generate">
      <input type="hidden" name="name" value="{{name}}">
      <input type="hidden" name="specialty" value="{{specialty}}">

      <label style="display:block;text-align:right;font-weight:700;margin-top:14px">اختر جهة التدريب</label>
      <select name="entity" required>
        <option value="" disabled selected>اختر جهة تدريبية...</option>
        {% for entity, count in entities %}
          <option value="{{entity}}">{{entity}} (متاح {{count}})</option>
        {% endfor %}
      </select>

      <button type="submit">طباعة خطاب التوجيه (PDF)</button>

      <div class="note">
        عند انتهاء فرص جهة معينة لن تظهر للاختيار.
      </div>
    </form>
  {% endif %}

  {% if error %}
    <div class="warn">{{error}}</div>
  {% endif %}
</div>

</body>
</html>
"""

SLOTS_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>الفرص حسب التخصص</title>
<style>
body{font-family:Arial;background:#f4f4f4;margin:0;padding:24px}
.card{max-width:1000px;margin:0 auto;background:#fff;border-radius:16px;padding:24px;box-shadow:0 10px 25px rgba(0,0,0,.08)}
h2{margin:0 0 14px 0}
table{width:100%;border-collapse:collapse}
th,td{padding:10px;border-bottom:1px solid #eee;text-align:right}
th{background:#fafafa}
a{color:#0b1a3a;text-decoration:none}
.small{color:#666;margin-top:10px}
</style>
</head>
<body>
<div class="card">
  <h2>الفرص المتاحة حسب التخصص</h2>
  <table>
    <thead>
      <tr><th>التخصص</th><th>إجمالي الفرص المتبقية</th></tr>
    </thead>
    <tbody>
      {% for spec, total in rows %}
      <tr><td>{{spec}}</td><td>{{total}}</td></tr>
      {% endfor %}
    </tbody>
  </table>
  <div class="small">
    <a href="/">رجوع للدخول</a> • <a href="/recalc">إعادة حساب الفرص من ملف الإكسل</a>
  </div>
</div>
</body>
</html>
"""

# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    try:
        if request.method == "POST":
            training_number = (request.form.get("training_number") or "").strip()
            mobile_last4 = (request.form.get("mobile_last4") or "").strip()

            df = load_students_df()

            col_training = find_col(df, ["رقم المتدرب", "StudentID", "ID"])
            col_mobile = find_col(df, ["رقم الجوال", "الجوال", "Mobile", "Phone"])
            col_name = find_col(df, ["اسم المتدرب", "الاسم", "Name"])

            try:
                col_spec = find_col(df, ["التخصص", "تخصص", "برنامج", "البرنامج", "Specialty", "Major", "Program"])
            except Exception:
                col_spec = "__ALL__"
                df[col_spec] = "عام"

            df[col_training] = df[col_training].astype(str).str.strip()
            df[col_mobile] = df[col_mobile].astype(str).apply(last4_digits)

            m = df[(df[col_training] == training_number) & (df[col_mobile] == mobile_last4)]
            if m.empty:
                error = "بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال."
            else:
                row = m.iloc[0]
                name = str(row[col_name]).strip()
                specialty = str(row[col_spec]).strip()
                return redirect(url_for("select", name=name, specialty=specialty))

    except Exception as e:
        error = f"حدث خطأ أثناء التحميل/التحقق:\n{e}"

    return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=error)

@app.route("/select", methods=["GET"])
def select():
    name = request.args.get("name", "").strip()
    specialty = request.args.get("specialty", "").strip()

    try:
        slots = load_slots_by_specialty()
        spec_slots = slots.get(specialty, {})
        entities = [(k, int(v)) for k, v in spec_slots.items() if int(v) > 0]
        entities.sort(key=lambda x: x[0])

        total_remaining = specialty_total_remaining(slots, specialty)
        no_entities = len(entities) == 0

        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            name=name,
            specialty=specialty,
            entities=entities,
            total_remaining=total_remaining,
            no_entities=no_entities,
            error=None
        )

    except Exception as e:
        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            name=name,
            specialty=specialty,
            entities=[],
            total_remaining=0,
            no_entities=True,
            error=f"خطأ في صفحة الاختيار:\n{e}"
        )

@app.route("/generate", methods=["POST"])
def generate():
    """
    يولد PDF ويخصم فرصة من الجهة داخل نفس التخصص.
    """
    try:
        name = (request.form.get("name") or "").strip()
        specialty = (request.form.get("specialty") or "").strip()
        entity = (request.form.get("entity") or "").strip()

        if not name or not specialty or not entity:
            return "بيانات ناقصة.", 400

        if not TEMPLATE_DOCX.exists():
            return f"ملف القالب غير موجود: {TEMPLATE_DOCX}", 500

        slots = load_slots_by_specialty()
        if specialty not in slots:
            return "التخصص غير موجود ضمن الفرص.", 400

        if int(slots[specialty].get(entity, 0)) <= 0:
            return "هذه الجهة لم تعد متاحة.", 400

        # خصم فرصة
        slots[specialty][entity] = int(slots[specialty][entity]) - 1

        # إذا انتهت تماماً، نتركها صفر (ستختفي من الاختيار)
        save_slots_by_specialty(slots)

        # بيانات القالب
        now_str = datetime.now().strftime("%Y-%m-%d")
        mapping = {
            "{{NAME}}": name,
            "{{ENTITY}}": entity,
            "{{SPECIALTY}}": specialty,
            "{{DATE}}": now_str,
        }

        pdf_id = uuid.uuid4().hex
        out_pdf = OUT_DIR / f"letter_{pdf_id}.pdf"
        render_docx_to_pdf(TEMPLATE_DOCX, out_pdf, mapping)

        return send_file(out_pdf, as_attachment=True, download_name="خطاب_التوجيه.pdf")

    except Exception as e:
        return (
            "حدث خطأ أثناء إنشاء PDF:\n"
            f"{e}\n\n"
            f"{traceback.format_exc()}",
            500,
        )

@app.route("/slots", methods=["GET"])
def slots_view():
    """
    يعرض كل تخصص وإجمالي الفرص المتبقية له.
    """
    try:
        slots = load_slots_by_specialty()
        rows = []
        for spec in sorted(slots.keys()):
            rows.append((spec, specialty_total_remaining(slots, spec)))
        return render_template_string(SLOTS_PAGE, rows=rows)
    except Exception as e:
        return f"خطأ: {e}", 500

@app.route("/recalc", methods=["GET"])
def recalc():
    """
    يعيد حساب الفرص من ملف الإكسل ويستبدل slots_by_specialty.json.
    استخدمه بعد ما تعدّل ملف students.xlsx
    """
    try:
        slots = calculate_slots_from_excel()
        save_slots_by_specialty(slots)
        return redirect(url_for("slots_view"))
    except Exception as e:
        return f"فشل إعادة الحساب:\n{e}", 500

# =========================
# Run
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

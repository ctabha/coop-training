
import os, json, shutil, subprocess, uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"
MAX_SLOTS_PER_ENTITY = 5

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "out"
OUT_DIR.mkdir(exist_ok=True)

EXCEL_PATH = DATA_DIR / "trainees.xlsx"
TEMPLATE_DOCX = DATA_DIR / "letter_template.docx"
SLOTS_PATH = DATA_DIR / "slots.json"
ASSIGNMENTS_PATH = DATA_DIR / "assignments.json"

# أسماء الأعمدة (حسب ملفك)
COL_TRAINEE_ID = "رقم المتدرب"
COL_NAME = "إسم المتدرب"
COL_MAJOR = "التخصص"
COL_PHONE = "رقم الجوال"
COL_ENTITY = "جهة التدريب "
COL_SUPERVISOR = "المدرب"
COL_COURSE_REF = "الرقم المرجعي"

app = Flask(__name__)

PAGE_HOME = """
<!doctype html><html lang="ar" dir="rtl"><head>
<meta charset="utf-8"/><title>{{title}}</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.wrap{max-width:760px;margin:40px auto;background:#fff;padding:24px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06)}
input,button{width:100%;padding:12px;margin:10px 0;border-radius:10px;border:1px solid #ddd;font-size:16px}
button{background:#111827;color:#fff;cursor:pointer;border:none}
.muted{color:#6b7280;font-size:14px}
.err{background:#fee2e2;color:#991b1b;padding:10px;border-radius:10px}
.ok{background:#dcfce7;color:#166534;padding:10px;border-radius:10px}
</style></head><body>
<div class="wrap">
<h2>{{title}}</h2>
<p class="muted">ادخل رقمك التدريبي ثم اطبع خطاب التوجيه (PDF). كل جهة لها {{max_slots}} فرص.</p>
{% if message %}<div class="{{'err' if is_error else 'ok'}}">{{message}}</div>{% endif %}
<form method="post" action="/check">
  <input name="trainee_id" placeholder="الرقم التدريبي" required />
  <input name="phone_last4" placeholder="آخر 4 أرقام من رقم الجوال" inputmode="numeric" maxlength="4" required />
  <button type="submit">دخول</button>
</form>
<hr/>
<p class="muted">للمشرف: صفحة الإحصاءات <b>/admin</b></p>
</div></body></html>
"""

PAGE_PICK = """
<!doctype html><html lang="ar" dir="rtl"><head>
<meta charset="utf-8"/><title>{{title}}</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.wrap{max-width:760px;margin:40px auto;background:#fff;padding:24px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06)}
select,input,button{width:100%;padding:12px;margin:10px 0;border-radius:10px;border:1px solid #ddd;font-size:16px}
button{background:#111827;color:#fff;cursor:pointer;border:none}
.muted{color:#6b7280;font-size:14px}
.info{background:#eff6ff;color:#1e3a8a;padding:10px;border-radius:10px}
</style></head><body>
<div class="wrap">
<h2>{{title}}</h2>
<div class="info">
  <div><b>الاسم:</b> {{name}}</div>
  <div><b>الرقم التدريبي:</b> {{trainee_id}}</div>
  <div><b>التخصص:</b> {{major}}</div>
  <div><b>الجوال:</b> {{phone}}</div>
</div>

<form method="post" action="/generate">
  <input type="hidden" name="trainee_id" value="{{trainee_id}}"/>
  <input type="hidden" name="phone_last4" value="{{phone_last4}}"/>
  <label class="muted">اختر جهة التدريب المتاحة:</label>
  <select name="entity" required>
    {% for e in entities %}
      <option value="{{e}}">{{e}} (المتبقي: {{remaining[e]}})</option>
    {% endfor %}
  </select>
  <button type="submit">طباعة خطاب التوجيه PDF</button>
</form>

<p class="muted">إذا لم تظهر جهة، فهذا يعني أن فرصها اكتملت ({{max_slots}}/{{max_slots}}) أو ليست ضمن جهاتك.</p>
<a class="muted" href="/">رجوع</a>
</div></body></html>
"""

PAGE_ADMIN = """
<!doctype html><html lang="ar" dir="rtl"><head>
<meta charset="utf-8"/><title>لوحة الإحصاءات</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.wrap{max-width:980px;margin:40px auto;background:#fff;padding:24px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,.06)}
table{width:100%;border-collapse:collapse}
th,td{border-bottom:1px solid #eee;padding:10px;text-align:right}
th{background:#fafafa}
.muted{color:#6b7280;font-size:14px}
.pill{display:inline-block;padding:4px 10px;border-radius:999px;background:#111827;color:#fff;font-size:12px}
</style></head><body>
<div class="wrap">
<h2>لوحة الإحصاءات</h2>
<p class="muted">كل جهة لها {{max_slots}} فرص — يتم الإغلاق تلقائيًا عند اكتمال العدد.</p>
<table>
<thead><tr><th>الجهة</th><th>المستخدم</th><th>المتبقي</th><th>الحالة</th></tr></thead>
<tbody>
{% for e, used in slots.items() %}
<tr>
<td>{{e}}</td>
<td><span class="pill">{{used}}</span></td>
<td>{{max_slots - used}}</td>
<td>{{"مغلقة" if used >= max_slots else "متاحة"}}</td>
</tr>
{% endfor %}
</tbody></table>
<hr/>
<p class="muted">إعادة ضبط (للمشرف): افتح <b>/admin/reset</b> لإرجاع العدّاد (يحذف الحجوزات).</p>
</div></body></html>
"""

def _load_excel() -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH)
    # تنظيف مسافات عمود جهة التدريب
    if COL_ENTITY in df.columns:
        df[COL_ENTITY] = df[COL_ENTITY].astype(str).str.strip()
    return df



def _digits(s: str) -> str:
    return "".join(ch for ch in str(s) if ch.isdigit())
def _load_json(path: Path, default):
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return default

def _save_json(path: Path, obj):
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")

def _get_trainee_row(df: pd.DataFrame, trainee_id: str):
    # مقارنة كنص بعد إزالة الفراغات
    tid = str(trainee_id).strip()
    s = df[COL_TRAINEE_ID].astype(str).str.strip()
    found = df[s == tid]
    if found.empty:
        return None
    return found.iloc[0]

def _remaining_by_entity(slots: dict) -> dict:
    return {e: max(0, MAX_SLOTS_PER_ENTITY - int(used)) for e, used in slots.items()}

def _available_entities_for_trainee(row, slots: dict):
    # في ملفك كل متدرب له "جهة التدريب" واحدة
    ent = str(row[COL_ENTITY]).strip()
    if not ent:
        return []
    remaining = _remaining_by_entity(slots)
    if remaining.get(ent, 0) <= 0:
        return []
    return [ent]

def _render_letter_docx(output_docx: Path, *, entity: str, trainee_id: str, name: str, major: str, phone: str, supervisor: str, course_ref: str):
    """
    ينسخ قالب DOCX ثم يعبّي الخانات داخل الجدول + سطر سعادة/ جهة التدريب
    """
    from docx import Document

    tmp_docx = OUT_DIR / f"tmp_{uuid.uuid4().hex}.docx"
    shutil.copy(TEMPLATE_DOCX, tmp_docx)

    d = Document(tmp_docx)

    # تعبئة الجدول (حسب ترتيب القالب المرفق)
    if d.tables:
        t = d.tables[0]
        # الصف 2 (index=1) يحتوي الخانات الفارغة: الرقم/الاسم/الرقم الأكاديمي/التخصص/جوال
        # الأعمدة 0..4
        try:
            t.cell(1,0).text = str(trainee_id)
            t.cell(1,1).text = str(name)
            t.cell(1,2).text = str(trainee_id)  # الرقم الأكاديمي
            t.cell(1,3).text = str(major)
            t.cell(1,4).text = str(phone)
        except Exception:
            pass

        # صف المشرف/الرقم المرجعي (القالب فيه دمج خلايا، لذلك نعبي كل الخلايا بنفس النص)
        try:
            # صف 3 (index=3) فارغ في القالب
            for c in range(len(t.rows[3].cells)):
                if c < 2:
                    t.cell(3,c).text = str(supervisor)
                else:
                    t.cell(3,c).text = str(course_ref)
        except Exception:
            pass

    # تعديل فقرة "سعادة /"
    for p in d.paragraphs:
        txt = p.text.strip()
        if txt.startswith("سعادة /") or txt == "سعادة /":
            p.text = f"سعادة / {entity}"
            break

    d.save(output_docx)

    # حذف المؤقت
    try:
        tmp_docx.unlink(missing_ok=True)
    except Exception:
        pass

def _convert_docx_to_pdf(input_docx: Path, output_pdf: Path):
    # تحويل عبر LibreOffice/soffice
    out_dir = output_pdf.parent
    out_dir.mkdir(exist_ok=True)
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", str(out_dir),
        str(input_docx)
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    # LibreOffice يسمي الملف بنفس اسم docx
    produced = out_dir / (input_docx.stem + ".pdf")
    if produced != output_pdf:
        produced.replace(output_pdf)

@app.get("/")
def home():
    return render_template_string(PAGE_HOME, title=APP_TITLE, max_slots=MAX_SLOTS_PER_ENTITY, message=None, is_error=False)

@app.post("/check")
def check():
    trainee_id = request.form.get("trainee_id","").strip()
    phone_last4 = request.form.get("phone_last4","").strip()
    df = _load_excel()
    row = _get_trainee_row(df, trainee_id)
    if row is None:
        return render_template_string(PAGE_HOME, title=APP_TITLE, max_slots=MAX_SLOTS_PER_ENTITY,
                                      message="لم يتم العثور على الرقم التدريبي في ملف المتدربين.", is_error=True)

    # تحقق من آخر 4 أرقام من الجوال
    phone_digits = _digits(row.get(COL_PHONE, ""))
    if len(phone_digits) < 4 or _digits(phone_last4)[-4:] != phone_digits[-4:]:
        return render_template_string(PAGE_HOME, title=APP_TITLE, max_slots=MAX_SLOTS_PER_ENTITY,
                                      message="رقم الجوال غير مطابق. أدخل آخر 4 أرقام من جوالك المسجل.", is_error=True)


    slots = _load_json(SLOTS_PATH, {})
    remaining = _remaining_by_entity(slots)
    entities = _available_entities_for_trainee(row, slots)

    if not entities:
        return render_template_string(PAGE_HOME, title=APP_TITLE, max_slots=MAX_SLOTS_PER_ENTITY,
                                      message="عذرًا، جهة تدريبك الحالية مغلقة لأن الفرص اكتملت.", is_error=True)

    return render_template_string(
        PAGE_PICK, title=APP_TITLE,
        trainee_id=str(trainee_id),
        name=str(row[COL_NAME]),
        major=str(row[COL_MAJOR]),
        phone=str(row[COL_PHONE]),
        phone_last4=_digits(phone_last4)[-4:],
        entities=entities,
        remaining=remaining,
        max_slots=MAX_SLOTS_PER_ENTITY
    )

@app.post("/generate")
def generate():
    trainee_id = request.form.get("trainee_id","").strip()
    phone_last4 = request.form.get("phone_last4","").strip()
    entity = request.form.get("entity","").strip()

    df = _load_excel()
    row = _get_trainee_row(df, trainee_id)
    if row is None:
        return redirect(url_for("home"))

    phone_digits = _digits(row.get(COL_PHONE, ""))
    if len(phone_digits) < 4 or _digits(phone_last4)[-4:] != phone_digits[-4:]:
        return redirect(url_for("home"))

    # تحميل العدّادات والحجوزات
    slots = _load_json(SLOTS_PATH, {})
    assignments = _load_json(ASSIGNMENTS_PATH, {})

    # إذا المتدرب سبق واختار جهة، لا نخصم مرة ثانية — نعيد نفس الـPDF
    if trainee_id in assignments:
        pdf_path = Path(assignments[trainee_id]["pdf"])
        if pdf_path.exists():
            return send_file(pdf_path, as_attachment=False, download_name=pdf_path.name)

    # تحقق من توفر الفرصة
    used = int(slots.get(entity, 0))
    if used >= MAX_SLOTS_PER_ENTITY:
        return render_template_string(PAGE_HOME, title=APP_TITLE, max_slots=MAX_SLOTS_PER_ENTITY,
                                      message="هذه الجهة مغلقة لأن فرصها اكتملت.", is_error=True)

    # إنشاء DOCX مخصص ثم تحويله إلى PDF
    out_docx = OUT_DIR / f"letter_{trainee_id}.docx"
    out_pdf = OUT_DIR / f"letter_{trainee_id}.pdf"

    _render_letter_docx(
        out_docx,
        entity=entity,
        trainee_id=str(trainee_id),
        name=str(row[COL_NAME]),
        major=str(row[COL_MAJOR]),
        phone=str(row[COL_PHONE]),
        supervisor=str(row.get(COL_SUPERVISOR, "")),
        course_ref=str(row.get(COL_COURSE_REF, "")),
    )
    _convert_docx_to_pdf(out_docx, out_pdf)

    # خصم الفرصة وتسجيل الحجز
    slots[entity] = used + 1
    assignments[trainee_id] = {"entity": entity, "pdf": str(out_pdf), "ts": datetime.now().isoformat()}

    _save_json(SLOTS_PATH, slots)
    _save_json(ASSIGNMENTS_PATH, assignments)

    return send_file(out_pdf, as_attachment=False, download_name=out_pdf.name)

@app.get("/admin")
def admin():
    slots = _load_json(SLOTS_PATH, {})
    # تأكيد ظهور الجهات حتى لو صفر
    df = _load_excel()
    entities = sorted(df[COL_ENTITY].astype(str).str.strip().dropna().unique())
    for e in entities:
        slots.setdefault(e, 0)
    _save_json(SLOTS_PATH, slots)
    return render_template_string(PAGE_ADMIN, slots=slots, max_slots=MAX_SLOTS_PER_ENTITY)

@app.get("/admin/reset")
def admin_reset():
    # إعادة ضبط العدّادات والحجوزات
    df = _load_excel()
    entities = sorted(df[COL_ENTITY].astype(str).str.strip().dropna().unique())
    slots = {e: 0 for e in entities}
    _save_json(SLOTS_PATH, slots)
    _save_json(ASSIGNMENTS_PATH, {})
    # حذف ملفات out
    for p in OUT_DIR.glob("letter_*.pdf"):
        try: p.unlink()
        except: pass
    for p in OUT_DIR.glob("letter_*.docx"):
        try: p.unlink()
        except: pass
    return "تمت إعادة الضبط بنجاح."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))

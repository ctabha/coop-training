import os
import json
from pathlib import Path

import pandas as pd
from flask import Flask, request, redirect, send_file, render_template_string

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm


# =========================
# إعدادات ومسارات
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

OUT_DIR.mkdir(exist_ok=True)

DATA_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"
SLOTS_FILE = DATA_DIR / "slots.json"


# =========================
# أدوات مساعدة
# =========================
def clean_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def save_json(path: Path, data):
    path.parent.mkdir(exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_students():
    # ملفك اسمه students.xlsx داخل data
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE)

    # توحيد أسماء الأعمدة حسب ملفك (الموجود بالصورة)
    # الأعمدة المتوقعة:
    # رقم المتدرب، اسم المتدرب، رقم الجوال، التخصص (و/أو برنامج)، جهة التدريب
    col_map = {}
    for c in df.columns:
        cc = clean_text(c)
        col_map[cc] = c

    required = ["رقم المتدرب", "اسم المتدرب", "رقم الجوال", "التخصص", "جهة التدريب"]
    for r in required:
        if r not in col_map:
            # إذا التخصص غير موجود لكن "برنامج" موجود نستخدمه بدل التخصص
            if r == "التخصص" and "برنامج" in col_map:
                col_map["التخصص"] = col_map["برنامج"]
                continue
            raise ValueError(f"العمود '{r}' غير موجود. الأعمدة الموجودة: {list(df.columns)}")

    # إعادة تسمية لتوحيد الاستخدام داخل الكود
    df = df.rename(
        columns={
            col_map["رقم المتدرب"]: "trainee_id",
            col_map["اسم المتدرب"]: "trainee_name",
            col_map["رقم الجوال"]: "phone",
            col_map["التخصص"]: "specialty",
            col_map["جهة التدريب"]: "entity",
        }
    )

    # تنظيف
    df["trainee_id"] = df["trainee_id"].astype(str).str.strip()
    df["phone"] = df["phone"].astype(str).str.strip()
    df["specialty"] = df["specialty"].astype(str).str.strip()
    df["entity"] = df["entity"].astype(str).str.strip()
    df["trainee_name"] = df["trainee_name"].astype(str).str.strip()

    return df


def build_slots_from_excel(df: pd.DataFrame):
    """
    يحسب الفرص من ملف الاكسل:
    كل جهة تدريب داخل نفس التخصص إذا تكررت = تزيد الفرص
    الناتج:
    slots = { "التخصص": { "جهة": عدد } }
    """
    slots = {}
    for _, row in df.iterrows():
        spec = clean_text(row.get("specialty"))
        ent = clean_text(row.get("entity"))
        if not spec or not ent:
            continue
        slots.setdefault(spec, {})
        slots[spec][ent] = slots[spec].get(ent, 0) + 1
    return slots


def ensure_slots_file(df):
    # إذا slots.json غير موجود، ننشئه من الاكسل
    if not SLOTS_FILE.exists():
        slots = build_slots_from_excel(df)
        save_json(SLOTS_FILE, slots)


def find_student(df, trainee_id, phone_last4):
    trainee_id = clean_text(trainee_id)
    phone_last4 = clean_text(phone_last4)

    if not trainee_id or not phone_last4:
        return None

    row = df[df["trainee_id"] == trainee_id]
    if row.empty:
        return None

    row = row.iloc[0].to_dict()
    phone = clean_text(row.get("phone"))
    if len(phone) < 4 or phone[-4:] != phone_last4:
        return "WRONG_PHONE"

    return row


def get_remaining_options(slots, specialty):
    specialty = clean_text(specialty)
    if specialty not in slots:
        return []
    opts = []
    for ent, cnt in slots[specialty].items():
        try:
            n = int(cnt)
        except:
            n = 0
        if n > 0:
            opts.append((ent, n))
    # ترتيب تنازلي حسب العدد
    opts.sort(key=lambda x: x[1], reverse=True)
    return opts


# =========================
# صفحات HTML
# =========================
PAGE_LOGIN = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.top-image{width:100%; height:25vh; overflow:hidden; background:#fff;}
.top-image img{width:100%; height:100%; object-fit:contain; display:block;}
.container{max-width:900px;margin:-30px auto 40px;background:#fff;border-radius:18px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:28px}
h1{text-align:center;margin:0 0 10px;font-size:44px}
p{text-align:center;color:#333}
.form-row{display:flex;gap:20px;margin-top:18px}
.form-row>div{flex:1}
label{display:block;margin:0 0 8px;font-weight:bold}
input{width:100%;padding:14px;border:1px solid #ddd;border-radius:14px;font-size:20px}
button{width:100%;margin-top:18px;padding:18px;border:0;border-radius:18px;background:#0b1730;color:#fff;font-size:22px;cursor:pointer}
.err{color:#c00;text-align:center;margin-top:14px;font-weight:bold}
.note{color:#666;text-align:center;margin-top:8px}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h1>{{title}}</h1>
  <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

  <form method="POST" action="/">
    <div class="form-row">
      <div>
        <label>الرقم التدريبي</label>
        <input name="trainee_id" placeholder="مثال: 444229747" required>
      </div>
      <div>
        <label>آخر 4 أرقام من الجوال</label>
        <input name="phone_last4" placeholder="مثال: 6101" required>
      </div>
    </div>
    <button type="submit">دخول</button>
  </form>

  {% if error %}
    <div class="err">{{error}}</div>
  {% endif %}
  {% if note %}
    <div class="note">{{note}}</div>
  {% endif %}
</div>
</body>
</html>
"""

PAGE_SELECT = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.top-image{width:100%; height:25vh; overflow:hidden; background:#fff;}
.top-image img{width:100%; height:100%; object-fit:contain; display:block;}
.container{max-width:900px;margin:-30px auto 40px;background:#fff;border-radius:18px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:28px}
h1{text-align:center;margin:0 0 6px;font-size:40px}
.meta{display:flex;gap:10px;justify-content:space-between;flex-wrap:wrap;margin:12px 0 16px}
.badge{background:#eef2ff;border-radius:999px;padding:10px 14px;font-weight:bold}
label{display:block;margin:10px 0 8px;font-weight:bold;font-size:18px}
select{width:100%;padding:14px;border:1px solid #ddd;border-radius:14px;font-size:18px}
button{width:100%;margin-top:18px;padding:18px;border:0;border-radius:18px;background:#0b1730;color:#fff;font-size:22px;cursor:pointer}
.err{color:#c00;text-align:center;margin-top:14px;font-weight:bold}
.ok{color:#0a7;text-align:center;margin-top:14px;font-weight:bold}
ul{margin:10px 0 0}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h1>{{title}}</h1>

  <div class="meta">
    <div class="badge">المتدرب: {{trainee_name}}</div>
    <div class="badge">رقم المتدرب: {{trainee_id}}</div>
    <div class="badge">التخصص: {{specialty}}</div>
  </div>

  {% if already %}
    <div class="ok">
      تم تسجيل اختيارك مسبقًا:<br>
      الجهة المختارة: <b>{{already_entity}}</b><br><br>
      <a href="/print?trainee_id={{trainee_id}}&phone_last4={{phone_last4}}">تحميل/طباعة خطاب التوجيه PDF</a>
    </div>
  {% else %}
    <form method="POST" action="/choose">
      <input type="hidden" name="trainee_id" value="{{trainee_id}}">
      <input type="hidden" name="phone_last4" value="{{phone_last4}}">
      <label>جهة التدريب المتاحة</label>
      <select name="entity" required>
        <option value="">اختر الجهة...</option>
        {% for ent, cnt in options %}
          <option value="{{ent}}">{{ent}} (متبقي: {{cnt}})</option>
        {% endfor %}
      </select>
      <button type="submit">حفظ الاختيار</button>
    </form>

    {% if no_options %}
      <div class="err">لا توجد جهات متاحة حاليًا لهذا التخصص.</div>
    {% endif %}
  {% endif %}

  <hr>
  <b>ملخص الفرص المتبقية داخل تخصصك:</b>
  <ul>
    {% for ent, cnt in options %}
      <li>{{ent}} : {{cnt}} فرصة</li>
    {% endfor %}
  </ul>

</div>
</body>
</html>
"""


# =========================
# Flask
# =========================
app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def login():
    error = ""
    note = ""

    df = load_students()
    ensure_slots_file(df)

    if request.method == "POST":
        trainee_id = request.form.get("trainee_id")
        phone_last4 = request.form.get("phone_last4")

        student = find_student(df, trainee_id, phone_last4)
        if student is None:
            error = "الرقم التدريبي غير موجود."
            return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error, note=note)

        if student == "WRONG_PHONE":
            error = "بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال."
            return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error, note=note)

        return redirect(f"/select?trainee_id={trainee_id}&phone_last4={phone_last4}")

    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error, note=note)


@app.route("/select")
def select_page():
    df = load_students()
    ensure_slots_file(df)

    trainee_id = request.args.get("trainee_id")
    phone_last4 = request.args.get("phone_last4")

    student = find_student(df, trainee_id, phone_last4)
    if student is None or student == "WRONG_PHONE":
        return redirect("/")

    trainee_name = student["trainee_name"]
    specialty = student["specialty"]

    slots = load_json(SLOTS_FILE, default={})
    assignments = load_json(ASSIGNMENTS_FILE, default={})

    already = str(trainee_id) in assignments
    already_entity = assignments.get(str(trainee_id), {}).get("entity", "")

    options = get_remaining_options(slots, specialty)
    no_options = len(options) == 0 and not already

    return render_template_string(
        PAGE_SELECT,
        title=APP_TITLE,
        trainee_id=trainee_id,
        phone_last4=phone_last4,
        trainee_name=trainee_name,
        specialty=specialty,
        options=options,
        no_options=no_options,
        already=already,
        already_entity=already_entity,
    )


@app.route("/choose", methods=["POST"])
def choose_entity():
    df = load_students()
    ensure_slots_file(df)

    trainee_id = request.form.get("trainee_id")
    phone_last4 = request.form.get("phone_last4")
    entity = clean_text(request.form.get("entity"))

    student = find_student(df, trainee_id, phone_last4)
    if student is None or student == "WRONG_PHONE":
        return redirect("/")

    specialty = clean_text(student["specialty"])
    trainee_name = clean_text(student["trainee_name"])

    slots = load_json(SLOTS_FILE, default={})
    assignments = load_json(ASSIGNMENTS_FILE, default={})

    # منع اختيار مرة ثانية
    if str(trainee_id) in assignments:
        return redirect(f"/select?trainee_id={trainee_id}&phone_last4={phone_last4}")

    # تحقق أن الجهة ضمن التخصص ومتاحة
    if specialty not in slots or entity not in slots[specialty]:
        return redirect(f"/select?trainee_id={trainee_id}&phone_last4={phone_last4}")

    remaining = int(slots[specialty].get(entity, 0))
    if remaining <= 0:
        return redirect(f"/select?trainee_id={trainee_id}&phone_last4={phone_last4}")

    # خصم فرصة
    slots[specialty][entity] = remaining - 1
    save_json(SLOTS_FILE, slots)

    # حفظ الاختيار
    assignments[str(trainee_id)] = {
        "trainee_id": str(trainee_id),
        "trainee_name": trainee_name,
        "specialty": specialty,
        "entity": entity,
    }
    save_json(ASSIGNMENTS_FILE, assignments)

    return redirect(f"/select?trainee_id={trainee_id}&phone_last4={phone_last4}")


@app.route("/print")
def print_letter():
    df = load_students()
    ensure_slots_file(df)

    assignments = load_json(ASSIGNMENTS_FILE, default={})

    trainee_id = request.args.get("trainee_id")
    phone_last4 = request.args.get("phone_last4")

    student = find_student(df, trainee_id, phone_last4)
    if student is None or student == "WRONG_PHONE":
        return redirect("/")

    assignment = assignments.get(str(trainee_id))
    if not assignment:
        return "لم يتم اختيار جهة تدريب بعد.", 400

    entity = assignment["entity"]
    trainee_name = assignment["trainee_name"]
    specialty = assignment["specialty"]

    pdf_path = OUT_DIR / f"letter_{trainee_id}.pdf"

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4

    y = height - 3 * cm
    c.setFont("Helvetica", 16)

    lines = [
        "خطاب توجيه تدريب تعاوني",
        "",
        f"اسم المتدرب: {trainee_name}",
        f"رقم المتدرب: {trainee_id}",
        f"التخصص: {specialty}",
        f"جهة التدريب: {entity}",
        "",
        "نتمنى لكم التوفيق والنجاح.",
    ]

    for line in lines:
        c.drawRightString(width - 2 * cm, y, line)
        y -= 1 * cm

    c.save()

    return send_file(pdf_path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

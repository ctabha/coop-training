import os
import json
import uuid
import shutil
import traceback
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, send_file

app = Flask(__name__)

# =========================
# إعدادات ومسارات
# =========================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

# اسم ملف الطلاب عندك في الصورة: students.xlsx
DATA_FILE_CANDIDATES = [
    DATA_DIR / "students.xlsx",
    DATA_DIR / "students.xlsx",
    BASE_DIR / "students.xlsx",
]

SLOTS_FILE = DATA_DIR / "slots_by_specialty.json"

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

# =========================
# أدوات مساعدة
# =========================
def last4_digits(x) -> str:
    s = "".join(ch for ch in str(x) if ch.isdigit())
    return s[-4:] if len(s) >= 4 else s

def norm_key(s: str) -> str:
    """تطبيع اسم العمود للمطابقة (حروف صغيرة + إزالة مسافات ورموز)"""
    s = str(s).strip().lower()
    for ch in [" ", "\t", "\n", "\r", "-", "_", ".", "ـ", "(", ")", "[", "]", "{", "}", ":", "؛", ",", "،", "/"]:
        s = s.replace(ch, "")
    return s

def pick_data_file() -> Path:
    for p in DATA_FILE_CANDIDATES:
        if p.exists():
            return p
    # إن لم يوجد
    raise FileNotFoundError(
        "لم أجد ملف الطلاب داخل data. المطلوب: data/students.xlsx"
    )

def load_students_df() -> pd.DataFrame:
    data_file = pick_data_file()
    df = pd.read_excel(data_file)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    """
    يبحث عن عمود موجود اعتمادًا على تطبيع الاسم.
    """
    norm_map = {norm_key(c): c for c in df.columns}

    for cand in candidates:
        k = norm_key(cand)
        if k in norm_map:
            return norm_map[k]

    # مطابقة جزئية (احتياط)
    for cand in candidates:
        kc = norm_key(cand)
        for real in df.columns:
            if kc in norm_key(real):
                return real

    raise KeyError(f"لم أجد عمود من: {candidates} | الأعمدة الموجودة: {list(df.columns)}")

def load_slots_by_specialty() -> dict:
    """
    يقرأ ملف الفرص، وإن لم يوجد يعيد حسابه من الاكسل.
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

def calculate_slots_from_excel() -> dict:
    """
    الفرص = عدد تكرار كل جهة داخل نفس التخصص في students.xlsx
    """
    df = load_students_df()

    col_entity = find_col(df, [
        "جهة التدريب", "الجهة", "جهة", "TrainingEntity", "Entity", "Company"
    ])

    # تخصص قد يكون "تخصص" أو "برنامج"
    try:
        col_spec = find_col(df, ["التخصص", "تخصص", "Specialty", "Major", "Program", "البرنامج", "برنامج"])
    except Exception:
        col_spec = "__ALL__"
        df[col_spec] = "عام"

    df[col_entity] = df[col_entity].astype(str).str.strip()
    df[col_spec] = df[col_spec].astype(str).str.strip()

    df = df[df[col_entity].notna() & (df[col_entity] != "")]

    slots = {}
    for spec, g in df.groupby(col_spec):
        counts = g[col_entity].value_counts().to_dict()
        slots[str(spec)] = {str(k): int(v) for k, v in counts.items()}

    return slots

# =========================
# صفحات HTML
# =========================
LOGIN_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f4f4f4;margin:0}
.top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
.top-image img{width:100%;height:100%;object-fit:cover;display:block}
.container{max-width:900px;margin:-40px auto 50px;background:#fff;padding:35px;border-radius:16px;text-align:center;box-shadow:0 10px 25px rgba(0,0,0,.08)}
.row{display:flex;gap:16px;justify-content:center;flex-wrap:wrap;margin-top:18px}
.field{flex:1;min-width:260px;text-align:right}
label{display:block;font-weight:700;margin:10px 0}
input{width:100%;padding:14px;border-radius:12px;border:1px solid #ddd;font-size:16px}
button{width:100%;background:#0b1a3a;color:#fff;padding:16px 24px;border:none;border-radius:14px;font-size:18px;margin-top:18px;cursor:pointer}
.error{color:#c00;margin-top:14px;font-weight:700;white-space:pre-wrap;text-align:right}
.small{color:#666;margin-top:10px}
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
.top-image{width:100%;height:25vh;overflow:hidden;background:#fff}
.top-image img{width:100%;height:100%;object-fit:cover;display:block}
.container{max-width:900px;margin:-40px auto 50px;background:#fff;padding:35px;border-radius:16px;text-align:center;box-shadow:0 10px 25px rgba(0,0,0,.08)}
select{width:100%;padding:14px;border-radius:12px;border:1px solid #ddd;font-size:16px;margin-top:10px}
button{width:100%;background:#0b1a3a;color:#fff;padding:16px 24px;border:none;border-radius:14px;font-size:18px;margin-top:18px;cursor:pointer}
.note{color:#666;margin-top:10px}
.warn{color:#c00;font-weight:700;margin-top:14px;white-space:pre-wrap;text-align:right}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h2 style="margin:0">مرحباً {{name}}</h2>
  <p class="note">تخصصك/برنامجك: <b>{{specialty}}</b></p>

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

      <button type="submit">إنشاء ملف (تجريبي)</button>
      <div class="note">ملاحظة: عند انتهاء فرص جهة معينة لن تظهر للاختيار.</div>
    </form>
  {% endif %}

  {% if error %}
    <div class="warn">{{error}}</div>
  {% endif %}
</div>

</body>
</html>
"""

# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])  # حتى لا يعطي Not Found
def login():
    error = None
    try:
        if request.method == "POST":
            training_number = (request.form.get("training_number") or "").strip()
            mobile_last4 = (request.form.get("mobile_last4") or "").strip()

            df = load_students_df()

            col_training = find_col(df, ["رقم المتدرب", "رقم_المتدرب", "StudentID", "ID", "رقم المتدرب/المتدربة"])
            col_mobile = find_col(df, ["رقم الجوال", "الجوال", "Mobile", "Phone"])
            col_name = find_col(df, ["اسم المتدرب", "اسم_المتدرب", "الاسم", "Name"])

            try:
                col_spec = find_col(df, ["التخصص", "تخصص", "Specialty", "Major", "Program", "البرنامج", "برنامج"])
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
    error = None

    try:
        slots = load_slots_by_specialty()
        spec_slots = slots.get(specialty, {})
        entities = [(k, int(v)) for k, v in spec_slots.items() if int(v) > 0]
        entities.sort(key=lambda x: x[0])
        no_entities = len(entities) == 0

        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            name=name,
            specialty=specialty,
            entities=entities,
            no_entities=no_entities,
            error=None
        )

    except Exception as e:
        error = f"خطأ في صفحة الاختيار:\n{e}"
        return render_template_string(
            SELECT_PAGE,
            title=APP_TITLE,
            name=name,
            specialty=specialty,
            entities=[],
            no_entities=True,
            error=error
        )

@app.route("/generate", methods=["POST"])
def generate():
    """
    (تجريبي) ينقص فرصة الجهة ويولد ملف نصي للتحميل.
    بعد ما تتأكد الأمور شغالة، أبدله لك لتوليد PDF الحقيقي.
    """
    try:
        name = (request.form.get("name") or "").strip()
        specialty = (request.form.get("specialty") or "").strip()
        entity = (request.form.get("entity") or "").strip()

        if not entity or not specialty:
            return "بيانات ناقصة."

        slots = load_slots_by_specialty()

        if specialty not in slots:
            return "تخصص غير موجود في الفرص."

        if slots[specialty].get(entity, 0) <= 0:
            return "هذه الجهة لم تعد متاحة."

        # نقص فرصة
        slots[specialty][entity] = int(slots[specialty][entity]) - 1
        save_slots_by_specialty(slots)

        file_id = uuid.uuid4().hex
        out_path = OUT_DIR / f"letter_{file_id}.txt"
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("خطاب توجيه - التدريب التعاوني (تجريبي)\n")
            f.write(f"الاسم: {name}\n")
            f.write(f"التخصص/البرنامج: {specialty}\n")
            f.write(f"الجهة المختارة: {entity}\n")
            f.write(f"التاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")

        return send_file(out_path, as_attachment=True)

    except Exception as e:
        # لا نخليها 500 بدون توضيح
        return f"حدث خطأ أثناء التوليد:\n{e}\n\n{traceback.format_exc()}", 500

# =========================
# تشغيل
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

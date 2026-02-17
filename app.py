import os
import json
import uuid
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, send_file

app = Flask(__name__)

# =========================
# Paths
# =========================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

DATA_FILE = DATA_DIR / "students.xlsx"
SLOTS_FILE = DATA_DIR / "slots_by_specialty.json"

OUT_DIR.mkdir(exist_ok=True)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

# =========================
# Helpers
# =========================
def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df

def find_col(df: pd.DataFrame, candidates: list[str]) -> str:
    """
    يرجع أول عمود موجود من قائمة مرشحين.
    """
    cols = set(df.columns)
    for c in candidates:
        if c in cols:
            return c
    # محاولة مطابقة جزئية (مثلاً "رقم الجوال" داخل "رقم الجوال للمتدرب")
    for c in candidates:
        for real in df.columns:
            if c in str(real):
                return real
    raise KeyError(f"لم أجد أي عمود من المرشحين: {candidates}. الأعمدة الموجودة: {list(df.columns)}")

def load_students() -> pd.DataFrame:
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {DATA_FILE}")
    df = pd.read_excel(DATA_FILE)
    df = normalize_cols(df)
    return df

def last4(s: str) -> str:
    s = "".join(ch for ch in str(s) if ch.isdigit())
    return s[-4:] if len(s) >= 4 else s

# =========================
# حساب الفرص حسب التخصص
# =========================
def calculate_slots_by_specialty() -> dict:
    """
    يبني قاموس:
    {
      "تخصص 1": {"جهة 1": 3, "جهة 2": 1},
      "تخصص 2": {"جهة 9": 5}
    }
    والفرص = عدد تكرار الجهة داخل ذلك التخصص في ملف الاكسل.
    """
    df = load_students()

    col_training_no = find_col(df, ["رقم المتدرب", "رقم_المتدرب", "StudentID", "ID"])
    col_mobile = find_col(df, ["رقم الجوال", "الجوال", "Mobile", "Phone"])
    col_name = find_col(df, ["اسم المتدرب", "اسم_المتدرب", "الاسم", "Name"])

    # التخصص قد يكون اسمه "تخصص" أو "برنامج"
    col_specialty = None
    try:
        col_specialty = find_col(df, ["التخصص", "تخصص", "Specialty", "Major", "Program", "البرنامج", "برنامج"])
    except Exception:
        # إذا ما لقيه نعتبر كل الطلاب ضمن تخصص واحد عام
        col_specialty = "__ALL__"
        df[col_specialty] = "عام"

    col_entity = find_col(df, ["جهة التدريب", "الجهة", "جهة", "TrainingEntity", "Company", "Entity"])

    # تنظيف
    df[col_specialty] = df[col_specialty].astype(str).str.strip()
    df[col_entity] = df[col_entity].astype(str).str.strip()

    # تجاهل الصفوف التي ليس لها جهة
    df = df[df[col_entity].notna() & (df[col_entity] != "")]

    slots_by_spec: dict[str, dict[str, int]] = {}
    for spec, g in df.groupby(col_specialty):
        counts = g[col_entity].value_counts().to_dict()
        slots_by_spec[str(spec)] = {str(k): int(v) for k, v in counts.items()}

    return slots_by_spec

def load_slots_by_specialty() -> dict:
    if not SLOTS_FILE.exists():
        slots = calculate_slots_by_specialty()
        with open(SLOTS_FILE, "w", encoding="utf-8") as f:
            json.dump(slots, f, ensure_ascii=False, indent=2)
        return slots

    with open(SLOTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def save_slots_by_specialty(slots: dict) -> None:
    with open(SLOTS_FILE, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

# =========================
# Pages (HTML)
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
.error{color:#c00;margin-top:14px;font-weight:700}
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
.warn{color:#c00;font-weight:700;margin-top:14px}
</style>
</head>
<body>

<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="container">
  <h2 style="margin:0">مرحباً {{name}}</h2>
  <p class="note">تخصصك: <b>{{specialty}}</b></p>

  {% if no_entities %}
    <div class="warn">لا توجد جهات تدريبية متاحة لهذا التخصص حالياً.</div>
  {% else %}
    <form method="post" action="/generate">
      <label style="display:block;text-align:right;font-weight:700;margin-top:14px">اختر جهة التدريب</label>
      <select name="entity" required>
        <option value="" disabled selected>اختر جهة تدريبية...</option>
        {% for entity, count in entities %}
          {% if count > 0 %}
            <option value="{{entity}}">{{entity}} (متاح {{count}})</option>
          {% endif %}
        {% endfor %}
      </select>

      <button type="submit">طباعة خطاب التوجيه PDF</button>
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
def login():
    error = None

    if request.method == "POST":
        training_number = (request.form.get("training_number") or "").strip()
        mobile_last4 = (request.form.get("mobile_last4") or "").strip()

        df = load_students()

        col_training_no = find_col(df, ["رقم المتدرب", "رقم_المتدرب", "StudentID", "ID"])
        col_mobile = find_col(df, ["رقم الجوال", "الجوال", "Mobile", "Phone"])
        col_name = find_col(df, ["اسم المتدرب", "اسم_المتدرب", "الاسم", "Name"])

        try:
            col_specialty = find_col(df, ["التخصص", "تخصص", "Specialty", "Major", "Program", "البرنامج", "برنامج"])
        except Exception:
            col_specialty = "__ALL__"
            df[col_specialty] = "عام"

        # مطابقة
        df[col_training_no] = df[col_training_no].astype(str).str.strip()
        df[col_mobile] = df[col_mobile].astype(str).apply(last4)

        m = df[
            (df[col_training_no] == str(training_number)) &
            (df[col_mobile] == str(mobile_last4))
        ]

        if m.empty:
            error = "بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال."
        else:
            row = m.iloc[0]
            name = str(row[col_name]).strip()
            specialty = str(row[col_specialty]).strip()
            return redirect(url_for("select", name=name, specialty=specialty))

    return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=error)

@app.route("/select", methods=["GET"])
def select():
    name = request.args.get("name", "").strip()
    specialty = request.args.get("specialty", "").strip()

    slots = load_slots_by_specialty()
    spec_slots = slots.get(specialty, {})

    # فقط الجهات التي ما زالت >0
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

@app.route("/generate", methods=["POST"])
def generate():
    # نقرأ اسم وتخصص من querystring للحفاظ على السياق (لو احتجته لاحقاً)
    # لكن هنا نعيد استخدام نفس specialty من referrer عادةً. الأفضل نمرره hidden لاحقاً،
    # للتبسيط سنعيد تحميله من "آخر اختيار" عبر request.referrer إن أمكن.
    entity = (request.form.get("entity") or "").strip()
    if not entity:
        return "لم يتم اختيار جهة."

    # نحاول استخراج specialty من referrer (select?specialty=...)
    specialty = ""
    ref = request.referrer or ""
    if "specialty=" in ref:
        try:
            specialty = ref.split("specialty=", 1)[1].split("&", 1)[0]
            specialty = specialty.replace("%20", " ").strip()
        except Exception:
            specialty = ""

    slots = load_slots_by_specialty()

    if specialty not in slots:
        # fallback: ابحث عن الجهة ضمن أي تخصص (لو referrer ما جاب التخصص)
        found_spec = None
        for sp, d in slots.items():
            if entity in d:
                found_spec = sp
                break
        if found_spec is None:
            return "تعذر تحديد التخصص/الجهة."
        specialty = found_spec

    if slots[specialty].get(entity, 0) <= 0:
        return "هذه الجهة لم تعد متاحة."

    # أنقص فرصة
    slots[specialty][entity] = int(slots[specialty][entity]) - 1
    save_slots_by_specialty(slots)

    # (مؤقت) توليد ملف بسيط — يمكنك لاحقاً استبداله بتوليد PDF الحقيقي
    file_id = uuid.uuid4().hex
    out_path = OUT_DIR / f"letter_{file_id}.txt"

    with open(out_path, "w", encoding="utf-8") as f:
        f.write("خطاب توجيه - التدريب التعاوني\n")
        f.write(f"التاريخ: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n")
        f.write(f"التخصص: {specialty}\n")
        f.write(f"الجهة المختارة: {entity}\n")

    return send_file(out_path, as_attachment=True)

# =========================
# Run
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

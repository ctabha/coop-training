import os
import json
import uuid
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, redirect, url_for, send_file

app = Flask(__name__)

# =========================
# إعداد المسارات
# =========================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "out"

DATA_FILE = DATA_DIR / "students.xlsx"
SLOTS_FILE = DATA_DIR / "slots.json"

OUT_DIR.mkdir(exist_ok=True)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

# =========================
# تحميل الطلاب
# =========================
def load_students():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.str.strip()
    return df

# =========================
# حساب الفرص من الاكسل
# =========================
def calculate_slots():
    df = load_students()
    counts = df["جهة التدريب"].value_counts().to_dict()
    return counts

# =========================
# تحميل أو إنشاء ملف الفرص
# =========================
def load_slots():
    if not SLOTS_FILE.exists():
        slots = calculate_slots()
        with open(SLOTS_FILE, "w", encoding="utf-8") as f:
            json.dump(slots, f, ensure_ascii=False, indent=2)
        return slots

    with open(SLOTS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

# =========================
# حفظ الفرص
# =========================
def save_slots(slots):
    with open(SLOTS_FILE, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

# =========================
# صفحة الدخول
# =========================
LOGIN_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f4f4f4;margin:0}
.container{max-width:700px;margin:80px auto;background:#fff;padding:30px;border-radius:12px;text-align:center}
input{width:90%;padding:12px;margin:10px;border-radius:8px;border:1px solid #ccc}
button{background:#0b1a3a;color:#fff;padding:12px 30px;border:none;border-radius:8px;font-size:16px}
.error{color:red;margin-top:15px}
</style>
</head>
<body>
<div class="container">
<h1>{{title}}</h1>
<p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>
<form method="post">
<input name="training_number" placeholder="الرقم التدريبي" required>
<input name="mobile_last4" placeholder="آخر 4 أرقام من الجوال" required>
<button type="submit">دخول</button>
</form>
{% if error %}
<div class="error">{{error}}</div>
{% endif %}
</div>
</body>
</html>
"""

# =========================
# صفحة اختيار الجهة
# =========================
SELECT_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f4f4f4;margin:0}
.container{max-width:800px;margin:50px auto;background:#fff;padding:30px;border-radius:12px;text-align:center}
select{width:90%;padding:12px;margin:15px;border-radius:8px}
button{background:#0b1a3a;color:#fff;padding:12px 30px;border:none;border-radius:8px;font-size:16px}
</style>
</head>
<body>
<div class="container">
<h2>مرحباً {{name}}</h2>
<form method="post" action="/generate">
<select name="entity" required>
{% for entity, count in slots.items() %}
{% if count > 0 %}
<option value="{{entity}}">{{entity}} (متاح {{count}})</option>
{% else %}
<option disabled>{{entity}} (غير متاح)</option>
{% endif %}
{% endfor %}
</select>
<br>
<button type="submit">طباعة خطاب التوجيه PDF</button>
</form>
</div>
</body>
</html>
"""

# =========================
# التحقق من الدخول
# =========================
@app.route("/", methods=["GET", "POST"])
def login():
    error = None

    if request.method == "POST":
        training_number = request.form.get("training_number").strip()
        mobile_last4 = request.form.get("mobile_last4").strip()

        df = load_students()

        student = df[
            (df["رقم المتدرب"].astype(str) == training_number) &
            (df["رقم الجوال"].astype(str).str[-4:] == mobile_last4)
        ]

        if not student.empty:
            name = student.iloc[0]["اسم المتدرب"]
            return redirect(url_for("select", name=name))
        else:
            error = "بيانات الدخول غير صحيحة."

    return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=error)

# =========================
# صفحة اختيار الجهة
# =========================
@app.route("/select")
def select():
    name = request.args.get("name")
    slots = load_slots()
    return render_template_string(SELECT_PAGE, title=APP_TITLE, name=name, slots=slots)

# =========================
# توليد الخطاب
# =========================
@app.route("/generate", methods=["POST"])
def generate():
    entity = request.form.get("entity")
    slots = load_slots()

    if slots.get(entity, 0) <= 0:
        return "الفرص انتهت لهذه الجهة."

    slots[entity] -= 1
    save_slots(slots)

    file_name = f"letter_{uuid.uuid4().hex}.txt"
    file_path = OUT_DIR / file_name

    with open(file_path, "w", encoding="utf-8") as f:
        f.write(f"خطاب توجيه إلى: {entity}\n")
        f.write(f"التاريخ: {datetime.now()}\n")

    return send_file(file_path, as_attachment=True)

# =========================
# تشغيل التطبيق
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

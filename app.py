import os
import uuid
import pandas as pd
import subprocess
from flask import Flask, request, redirect, url_for, render_template_string, send_file

app = Flask(__name__)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data", "students.xlsx")
STATIC_FOLDER = os.path.join(BASE_DIR, "static")

# ----------------------------
# قراءة ملف الطلاب
# ----------------------------
def load_students():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.str.strip()
    return df

# ----------------------------
# حساب الفرص المتاحة حسب تكرار الجهة
# ----------------------------
def calculate_available_slots():
    df = load_students()
    counts = df["جهة التدريب"].value_counts().to_dict()
    return counts

# ----------------------------
# صفحة تسجيل الدخول
# ----------------------------
LOGIN_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.top-image{width:100%;height:25vh;overflow:hidden}
.top-image img{width:100%;height:100%;object-fit:cover}
.container{width:600px;margin:30px auto;background:white;padding:30px;border-radius:15px;text-align:center}
input{width:90%;padding:12px;margin:10px;border-radius:10px;border:1px solid #ccc;font-size:16px}
button{width:95%;padding:14px;background:#0d1b3d;color:white;border:none;border-radius:12px;font-size:18px;cursor:pointer}
.error{color:red;margin-top:15px}
</style>
</head>
<body>

<div class="top-image">
<img src="/static/header.jpg">
</div>

<div class="container">
<h2>{{title}}</h2>
<p>يرجى إدخال الرقم التدريبي وآخر 4 أرقام من الجوال</p>

<form method="post">
<input type="text" name="student_id" placeholder="الرقم التدريبي" required>
<input type="text" name="phone_last4" placeholder="آخر 4 أرقام من الجوال" required>
<button type="submit">دخول</button>
</form>

{% if error %}
<div class="error">{{error}}</div>
{% endif %}
</div>

</body>
</html>
"""

# ----------------------------
# صفحة الفرص المتاحة
# ----------------------------
SLOTS_PAGE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>الفرص المتاحة</title>
<style>
body{font-family:Arial;background:#f7f7f7;margin:0}
.top-image{width:100%;height:25vh;overflow:hidden}
.top-image img{width:100%;height:100%;object-fit:cover}
.container{width:700px;margin:30px auto;background:white;padding:30px;border-radius:15px}
button{padding:12px 20px;background:#0d1b3d;color:white;border:none;border-radius:10px;cursor:pointer}
</style>
</head>
<body>

<div class="top-image">
<img src="/static/header.jpg">
</div>

<div class="container">
<h2>الفرص المتاحة حسب الجهات</h2>

<ul>
{% for org, count in slots.items() %}
<li>{{org}} : {{count}} فرصة</li>
{% endfor %}
</ul>

<br>
<a href="{{url_for('generate_letter', student_id=student_id)}}">
<button>طباعة خطاب التوجيه PDF</button>
</a>

</div>
</body>
</html>
"""

# ----------------------------
# إنشاء ملف PDF
# ----------------------------
def generate_pdf(student):
    html_content = f"""
    <html dir="rtl">
    <body style="font-family:Arial">
    <h2>خطاب توجيه تدريب تعاوني</h2>
    <p>الاسم: {student['اسم المتدرب']}</p>
    <p>الرقم التدريبي: {student['رقم المتدرب']}</p>
    <p>جهة التدريب: {student['جهة التدريب']}</p>
    </body>
    </html>
    """

    file_id = str(uuid.uuid4())
    html_path = os.path.join(BASE_DIR, f"{file_id}.html")
    pdf_path = os.path.join(BASE_DIR, f"{file_id}.pdf")

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)

    subprocess.run(["libreoffice", "--headless", "--convert-to", "pdf", html_path, "--outdir", BASE_DIR])

    os.remove(html_path)

    return pdf_path

# ----------------------------
# المسارات
# ----------------------------
@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        student_id = request.form["student_id"]
        phone_last4 = request.form["phone_last4"]

        df = load_students()

        student = df[
            (df["رقم المتدرب"].astype(str) == student_id) &
            (df["رقم الجوال"].astype(str).str[-4:] == phone_last4)
        ]

        if not student.empty:
            return redirect(url_for("slots", student_id=student_id))
        else:
            error = "بيانات الدخول غير صحيحة"

    return render_template_string(LOGIN_PAGE, title=APP_TITLE, error=error)


@app.route("/slots/<student_id>")
def slots(student_id):
    slots = calculate_available_slots()
    return render_template_string(SLOTS_PAGE, slots=slots, student_id=student_id)


@app.route("/generate/<student_id>")
def generate_letter(student_id):
    df = load_students()
    student = df[df["رقم المتدرب"].astype(str) == student_id]

    if student.empty:
        return "الطالب غير موجود"

    pdf_path = generate_pdf(student.iloc[0])
    return send_file(pdf_path, as_attachment=True)


# ----------------------------
# تشغيل التطبيق
# ----------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

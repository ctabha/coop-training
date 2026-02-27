from flask import Flask, request, send_file, render_template_string
import pandas as pd
from docxtpl import DocxTemplate
from io import BytesIO
import os

app = Flask(__name__)

DATA_FILE = "data/students.xlsx"
TEMPLATE_FILE = "data/letter_template.docx"

HTML = """
<!DOCTYPE html>
<html lang="ar">
<head>
<meta charset="UTF-8">
<title>بوابة خطاب التوجيه - التدريب التعاوني</title>

<style>
html, body {
    height: 100%;
    margin: 0;
    font-family: Tahoma, Arial;
}

body {
    background: url("/static/header.jpg") no-repeat center center fixed;
    background-size: cover;
    direction: rtl;
}

.page-wrap {
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
}

.card {
    width: 900px;
    background: rgba(255,255,255,0.95);
    padding: 40px;
    border-radius: 18px;
    text-align: center;
}

input {
    width: 100%;
    height: 55px;
    font-size: 18px;
    margin-bottom: 18px;
    padding: 0 15px;
    border: 1px solid #ccc;
    border-radius: 6px;
}

button {
    padding: 12px 30px;
    font-size: 18px;
    background: #0a1a33;
    color: white;
    border: none;
    border-radius: 8px;
    cursor: pointer;
}
</style>
</head>

<body>
<div class="page-wrap">
  <div class="card">

    <h2>بوابة خطاب التوجيه - التدريب التعاوني</h2>

    <form method="POST">
      <input name="trainee_id" placeholder="الرقم التدريبي / رقم المتدرب" required>
      <input name="phone_last4" placeholder="آخر 4 أرقام من الجوال" required>
      <button type="submit">دخول</button>
    </form>

    <div style="color:red; margin-top:15px;">{{message}}</div>

  </div>
</div>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    message = ""

    if request.method == "POST":
        trainee_id = request.form.get("trainee_id")
        phone_last4 = request.form.get("phone_last4")

        if not os.path.exists(DATA_FILE):
            return render_template_string(HTML, message="ملف الطلاب غير موجود")

        df = pd.read_excel(DATA_FILE)

        required_cols = ["رقم المتدرب", "اسم المتدرب", "رقم الجوال"]

        for col in required_cols:
            if col not in df.columns:
                return render_template_string(
                    HTML,
                    message="أحد الأعمدة الأساسية غير موجود في Excel"
                )

        df["رقم المتدرب"] = df["رقم المتدرب"].astype(str)
        df["رقم الجوال"] = df["رقم الجوال"].astype(str)

        student = df[df["رقم المتدرب"] == str(trainee_id)]

        if student.empty:
            return render_template_string(HTML, message="رقم المتدرب غير صحيح")

        student = student.iloc[0]

        if not student["رقم الجوال"].endswith(str(phone_last4)):
            return render_template_string(HTML, message="آخر 4 أرقام غير صحيحة")

        if not os.path.exists(TEMPLATE_FILE):
            return render_template_string(HTML, message="قالب الخطاب غير موجود")

        doc = DocxTemplate(TEMPLATE_FILE)

        context = {
            "TraineeName": student["اسم المتدرب"],
            "AcademicID": student["رقم المتدرب"],
            "Phone": student["رقم الجوال"],
            "Company": student.get("جهة التدريب", ""),
            "Specialty": student.get("التخصص", ""),
            "LetterNo": student.get("الرقم المرجعي", ""),
        }

        doc.render(context)

        file_stream = BytesIO()
        doc.save(file_stream)
        file_stream.seek(0)

        return send_file(
            file_stream,
            as_attachment=True,
            download_name="خطاب_توجيه.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    return render_template_string(HTML, message=message)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

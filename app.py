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
<html dir="rtl">
<head>
<meta charset="UTF-8">
<title>بوابة خطاب التوجيه</title>
</head>
<body style="text-align:center">
<h2>بوابة خطاب التوجيه - التدريب التعاوني</h2>
<form method="POST">
<input name="trainee_id" placeholder="الرقم التدريبي" required><br><br>
<input name="phone_last4" placeholder="آخر 4 أرقام من الجوال" required><br><br>
<button type="submit">دخول</button>
</form>
<p style="color:red;">{{message}}</p>
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def index():
    message = ""

    if request.method == "POST":
        trainee_id = request.form["trainee_id"]
        last4 = request.form["phone_last4"]

        if not os.path.exists(DATA_FILE):
            return render_template_string(HTML, message="ملف الطلاب غير موجود")

        df = pd.read_excel(DATA_FILE)
        df["رقم المتدرب"] = df["رقم المتدرب"].astype(str)
        df["رقم الجوال"] = df["رقم الجوال"].astype(str)

        student = df[df["رقم المتدرب"] == trainee_id]

        if student.empty:
            return render_template_string(HTML, message="لم يتم العثور على المتدرب")

        student = student.iloc[0]

        if not student["رقم الجوال"].endswith(last4):
            return render_template_string(HTML, message="آخر 4 أرقام غير صحيحة")

        if not os.path.exists(TEMPLATE_FILE):
            return render_template_string(HTML, message="قالب الخطاب غير موجود")

        doc = DocxTemplate(TEMPLATE_FILE)

        context = {
            "trainee_name": student["اسم المتدرب"],
            "trainee_id": student["رقم المتدرب"],
            "phone": student["رقم الجوال"],
            "college_supervisor": student.get("المدرب", ""),
            "training_entity": student.get("جهة التدريب", ""),
            "course_ref": student.get("الرقم المرجعي", "")
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

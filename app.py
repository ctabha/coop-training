from flask import Flask, request, render_template_string, send_file
import pandas as pd
import os
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime

app = Flask(__name__)

DATA_FILE = "data/students.xlsx"
TEMPLATE_FILE = "data/letter_template.docx"


def replace_text_in_doc(doc, replacements):
    for paragraph in doc.paragraphs:
        for key, value in replacements.items():
            if key in paragraph.text:
                paragraph.text = paragraph.text.replace(key, str(value))

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in replacements.items():
                    if key in cell.text:
                        cell.text = cell.text.replace(key, str(value))


@app.route("/", methods=["GET", "POST"])
def index():
    message = ""

    if request.method == "POST":
        training_number = request.form.get("training_number")
        last4 = request.form.get("last4")

        if not os.path.exists(DATA_FILE):
            message = "ملف الطلاب غير موجود"
            return render_template_string(HTML_PAGE, message=message)

        df = pd.read_excel(DATA_FILE)

        if "رقم المتدرب" not in df.columns or "رقم الجوال" not in df.columns:
            message = "الأعمدة المطلوبة غير موجودة في ملف الطلاب"
            return render_template_string(HTML_PAGE, message=message)

        student = df[
            (df["رقم المتدرب"].astype(str) == str(training_number)) &
            (df["رقم الجوال"].astype(str).str[-4:] == str(last4))
        ]

        if student.empty:
            message = "لم يتم العثور على بيانات لهذا المتدرب"
            return render_template_string(HTML_PAGE, message=message)

        student = student.iloc[0]

        if not os.path.exists(TEMPLATE_FILE):
            message = "قالب الخطاب غير موجود"
            return render_template_string(HTML_PAGE, message=message)

        doc = Document(TEMPLATE_FILE)

        replacements = {
            "{{اسم_المتدرب}}": student.get("اسم المتدرب", ""),
            "{{رقم_المتدرب}}": student.get("رقم المتدرب", ""),
            "{{القسم}}": student.get("القسم", ""),
            "{{التخصص}}": student.get("التخصص", ""),
            "{{الجهة}}": student.get("جهة التدريب", ""),
            "{{التاريخ}}": datetime.today().strftime("%Y-%m-%d")
        }

        replace_text_in_doc(doc, replacements)

        output_path = f"letter_{training_number}.docx"
        doc.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template_string(HTML_PAGE, message=message)


HTML_PAGE = """
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="UTF-8">
<title>بوابة خطاب التوجيه - التدريب التعاوني</title>
<style>
body { font-family: Arial; background:#f2f2f2; text-align:center; }
.box { background:white; width:60%; margin:50px auto; padding:30px; border-radius:10px; }
input { padding:10px; margin:10px; width:60%; }
button { padding:10px 30px; background:#0b1d3a; color:white; border:none; border-radius:5px; }
.error { color:red; margin-top:20px; }
</style>
</head>
<body>
<div class="box">
<h2>بوابة خطاب التوجيه - التدريب التعاوني</h2>
<form method="POST">
<input type="text" name="training_number" placeholder="الرقم التدريبي" required><br>
<input type="text" name="last4" placeholder="آخر 4 أرقام من الجوال" required><br>
<button type="submit">دخول</button>
</form>
<div class="error">{{message}}</div>
</div>
</body>
</html>
"""

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

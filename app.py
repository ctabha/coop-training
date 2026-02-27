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
<html dir="rtl" lang="ar">
<head>
<meta charset="UTF-8">
<title>بوابة خطاب التوجيه</title>
<style>
body{font-family:Arial; text-align:center; padding:40px;}
input{width:420px; padding:14px; margin:10px; font-size:18px; text-align:right;}
button{padding:12px 30px; font-size:18px; cursor:pointer;}
.msg{color:red; margin-top:15px; font-size:18px;}
</style>
</head>
<body>
<h2>بوابة خطاب التوجيه - التدريب التعاوني</h2>
<form method="POST">
  <input name="trainee_id" placeholder="الرقم التدريبي / رقم المتدرب" required>
  <br>
  <input name="phone_last4" placeholder="آخر 4 أرقام من الجوال" required>
  <br>
  <button type="submit">دخول</button>
</form>
<div class="msg">{{message}}</div>
</body>
</html>
"""

def norm(s: str) -> str:
    # توحيد الاسم لتجنب مشاكل المسافات والرموز
    return "".join(str(s).strip().replace("\ufeff", "").split())

def find_col(cols, candidates):
    """
    cols: list of columns
    candidates: list of keywords we accept (normalized match)
    """
    ncols = {norm(c): c for c in cols}
    # تطابق مباشر
    for cand in candidates:
        if cand in ncols:
            return ncols[cand]
    # تطابق جزئي (contains)
    for cand in candidates:
        for n, original in ncols.items():
            if cand in n:
                return original
    return None

@app.route("/", methods=["GET", "POST"])
def index():
    message = ""

    try:
        if request.method == "POST":
            trainee_id = request.form.get("trainee_id", "").strip()
            last4 = request.form.get("phone_last4", "").strip()

            if not os.path.exists(DATA_FILE):
                return render_template_string(HTML, message="❌ ملف الطلاب غير موجود داخل data/students.xlsx")

            df = pd.read_excel(DATA_FILE)
            df.columns = [str(c) for c in df.columns]  # تأكيد أنها نصوص

            # ابحث عن الأعمدة الأساسية حتى لو كان اسمها مختلف
            col_id = find_col(df.columns, [
                norm("رقم المتدرب"),
                norm("الرقم التدريبي"),
                norm("الرقم_التدريبي"),
                norm("trainee_id"),
            ])

            col_name = find_col(df.columns, [
                norm("اسم المتدرب"),
                norm("اسم_المتدرب"),
                norm("trainee_name"),
            ])

            col_phone = find_col(df.columns, [
                norm("رقم الجوال"),
                norm("رقم_الجوال"),
                norm("الجوال"),
                norm("الهاتف"),
                norm("phone"),
            ])

            # أعمدة اختيارية
            col_supervisor = find_col(df.columns, [norm("المدرب"), norm("المشرف"), norm("college_supervisor")])
            col_entity = find_col(df.columns, [norm("جهة التدريب"), norm("جهة_التدريب"), norm("training_entity")])
            col_course_ref = find_col(df.columns, [norm("الرقم المرجعي"), norm("الرقمالمرجعي"), norm("course_ref")])

            # لو ناقص عمود أساسي نعرض رسالة بدل 500
            missing = []
            if not col_id: missing.append("رقم المتدرب/الرقم التدريبي")
            if not col_name: missing.append("اسم المتدرب")
            if not col_phone: missing.append("رقم الجوال")

            if missing:
                return render_template_string(
                    HTML,
                    message="❌ الأعمدة الأساسية غير موجودة في Excel: " + " ، ".join(missing) +
                            "<br>✅ الأعمدة الموجودة عندك: " + " | ".join(df.columns.astype(str).tolist())
                )

            # تحويل القيم نص عشان المطابقة
            df[col_id] = df[col_id].astype(str).str.strip()
            df[col_phone] = df[col_phone].astype(str).str.strip()

            student = df[df[col_id] == str(trainee_id)]
            if student.empty:
                return render_template_string(HTML, message="❌ لم يتم العثور على متدرب بهذا الرقم")

            student = student.iloc[0]
            phone_full = str(student[col_phone])

            if not phone_full.endswith(str(last4)):
                return render_template_string(HTML, message="❌ آخر 4 أرقام من الجوال غير صحيحة")

            if not os.path.exists(TEMPLATE_FILE):
                return render_template_string(HTML, message="❌ قالب الخطاب غير موجود: data/letter_template.docx")

            doc = DocxTemplate(TEMPLATE_FILE)

            context = {
                "trainee_name": str(student[col_name]),
                "trainee_id": str(student[col_id]),
                "phone": phone_full,
                "college_supervisor": str(student[col_supervisor]) if col_supervisor else "",
                "training_entity": str(student[col_entity]) if col_entity else "",
                "course_ref": str(student[col_course_ref]) if col_course_ref else "",
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

    except Exception as e:
        # بدل Internal Server Error نعطيك السبب على الصفحة
        return render_template_string(HTML, message=f"❌ خطأ داخل التطبيق: {e}")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

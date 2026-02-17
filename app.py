import os
from pathlib import Path

from flask import Flask, request, render_template_string, send_file, url_for

app = Flask(__name__)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
OUT_DIR = BASE_DIR / "out"
OUT_DIR.mkdir(exist_ok=True)

# =========================
# الصفحة الرئيسية (واجهة)
# =========================
PAGE_HOME = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }}</title>

  <style>
    body{
      font-family: Arial, sans-serif;
      background:#f7f7f7;
      margin:0;
      padding:0;
      direction: rtl;
    }

    /* صورة الهيدر: ربع الصفحة من الأعلى */
    .top-image{
      width:100%;
      height:25vh;          /* ربع الصفحة تقريباً */
      overflow:hidden;
      background:#fff;
    }
    .top-image img{
      width:100%;
      height:100%;
      object-fit:cover;     /* تملى العرض والارتفاع بدون تمدد */
      display:block;
    }

    /* محتوى الصفحة */
    .container{
      max-width: 900px;
      margin: 30px auto;
      background:#fff;
      padding: 28px;
      border-radius: 14px;
      box-shadow: 0 8px 24px rgba(0,0,0,.08);
      text-align:center;
    }

    h1{
      margin: 0 0 12px;
      font-size: 40px;
      font-weight: 800;
    }

    p{
      margin: 0 0 20px;
      color:#333;
      font-size: 18px;
    }

    .btn{
      display:inline-block;
      padding: 14px 22px;
      border:0;
      border-radius: 10px;
      background:#111827;
      color:#fff;
      cursor:pointer;
      font-size: 16px;
      text-decoration:none;
    }
    .btn:hover{ opacity:.92; }

    .note{
      margin-top:14px;
      font-size: 14px;
      color:#666;
    }
  </style>
</head>

<body>
  <!-- الهيدر -->
  <div class="top-image">
    <img src="{{ url_for('static', filename='header.jpg') }}" alt="Header Image">
  </div>

  <!-- المحتوى -->
  <div class="container">
    <h1>{{ title }}</h1>
    <p>مرحباً بك في نظام إصدار خطاب التدريب التعاوني</p>

    <!-- زر تجريبي (اختياري) -->
    <form method="post" action="/generate">
      <button class="btn" type="submit">طباعة خطاب التوجيه PDF</button>
    </form>

    <div class="note">
      إذا ظهر خطأ عند الطباعة: افتح Logs في Render ثم Live tail وأرسل لي آخر سطرين من الخطأ.
    </div>
  </div>
</body>
</html>
"""

@app.get("/")
def home():
    return render_template_string(PAGE_HOME, title=APP_TITLE)


# =========================
# مسار تجريبي للطباعة PDF
# (حالياً يعطي PDF بسيط)
# =========================
@app.post("/generate")
def generate():
    # PDF بسيط للتأكد أن زر الطباعة يعمل (بدون docx)
    pdf_path = OUT_DIR / "sample.pdf"

    # نص PDF بسيط جداً (بايتات) — كحل سريع للاختبار
    # إذا عندك نظام docx->pdf القديم، قلّي وأرجّعه لك داخل هذا الملف.
    pdf_bytes = b"%PDF-1.4\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>\nendobj\n4 0 obj\n<< /Length 44 >>\nstream\nBT /F1 18 Tf 50 780 Td (PDF is working!) Tj ET\nendstream\nendobj\nxref\n0 5\n0000000000 65535 f \n0000000010 00000 n \n0000000056 00000 n \n0000000111 00000 n \n0000000216 00000 n \ntrailer\n<< /Size 5 /Root 1 0 R >>\nstartxref\n310\n%%EOF\n"
    pdf_path.write_bytes(pdf_bytes)

    return send_file(pdf_path, as_attachment=True, download_name="letter.pdf")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

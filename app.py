import os
from flask import Flask, render_template_string

app = Flask(__name__)

APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

PAGE_HOME = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<title>{{title}}</title>

<style>
body{
    font-family:Arial;
    background:#f7f7f7;
    margin:0;
}

/* الصورة العلوية */
.top-image{
    width:100%;
    height:25vh;   /* ربع الشاشة */
    overflow:hidden;
}

.top-image img{
    width:100%;
    height:100%;
    object-fit:cover;
}

/* الحاوية */
.container{
    max-width:900px;
    margin:30px auto;
    background:white;
    padding:30px;
    border-radius:12px;
    box-shadow:0 4px 10px rgba(0,0,0,0.1);
    text-align:center;
}
</style>

</head>

<body>

<div class="top-image">
    <img src="/static/header.jpg" alt="Header Image">
</div>

<div class="container">
    <h1>{{title}}</h1>
    <p>مرحباً بك في نظام إصدار خطاب التدريب التعاوني</p>
</div>

</body>
</html>
"""

@app.route("/")
def home():
    return render_template_string(PAGE_HOME, title=APP_TITLE)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

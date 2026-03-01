from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

DATA_PATH = os.path.join("data", "students.xlsx")

def normalize_phone(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # إزالة .0 لو كانت أرقام من إكسل
    if s.endswith(".0"):
        s = s[:-2]
    # إزالة المسافات والرموز
    s = "".join(ch for ch in s if ch.isdigit())
    return s

@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    trainee = None

    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال"
            return render_template("index.html", error=error)

        if not os.path.exists(DATA_PATH):
            error = f"ملف الطلاب غير موجود: {DATA_PATH}"
            return render_template("index.html", error=error)

        df = pd.read_excel(DATA_PATH)

        # تأكد من وجود الأعمدة الأساسية
        required_cols = ["رقم المتدرب", "رقم الجوال"]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            error = f"الأعمدة التالية غير موجودة في Excel: {', '.join(missing)}"
            return render_template("index.html", error=error)

        # فلترة المتدرب
        df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
        df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)

        matches = df[df["رقم المتدرب"] == str(trainee_id)]

        if matches.empty:
            error = "لم يتم العثور على رقم المتدرب"
            return render_template("index.html", error=error)

        # أخذ أول نتيجة كقاموس (وهذا هو حل الخطأ)
        trainee_list = matches.to_dict("records")
        trainee = trainee_list[0]

        phone = normalize_phone(trainee.get("رقم الجوال", ""))
        if len(phone) < 4 or phone[-4:] != last4:
            error = "آخر 4 أرقام من الجوال غير صحيحة"
            return render_template("index.html", error=error)

        return render_template("index.html", trainee=trainee)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)

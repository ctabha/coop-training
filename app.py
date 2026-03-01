from flask import Flask, render_template, request
import pandas as pd
import os
import re

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(BASE_DIR, "data", "students.xlsx")

# تنظيف رقم الجوال واستخراج آخر 4 أرقام
def get_last4(phone):
    digits = re.sub(r"\D", "", str(phone))
    return digits[-4:] if len(digits) >= 4 else ""

# تحميل ملف الطلاب
def load_students():
    if not os.path.exists(DATA_PATH):
        raise FileNotFoundError("ملف students.xlsx غير موجود داخل مجلد data")

    df = pd.read_excel(DATA_PATH, dtype=str).fillna("")
    df.columns = [str(col).strip() for col in df.columns]

    required_cols = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب"]
    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        raise ValueError(
            "الأعمدة التالية غير موجودة في Excel:\n"
            + " | ".join(missing)
            + "\n\nالأعمدة الموجودة:\n"
            + " | ".join(df.columns)
        )

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["__last4__"] = df["رقم الجوال"].apply(get_last4)

    return df

@app.route("/", methods=["GET", "POST"])
def index():
    trainee = None
    error = ""

    if request.method == "POST":
        trainee_id = request.form.get("trainee_id", "").strip()
        phone4 = request.form.get("phone4", "").strip()

        try:
            df = load_students()

            row = df[
                (df["رقم المتدرب"] == trainee_id) &
                (df["__last4__"] == phone4)
            ]

            if row.empty:
                error = "الرقم التدريبي أو آخر 4 أرقام من الجوال غير صحيحة."
                return render_template("index.html", trainee=None, error=error)

            # ✅ حل المشكلة: نأخذ أول سجل كقاموس
            trainee = row.iloc[0].to_dict()

            return render_template("index.html", trainee=trainee, error="")

        except Exception as e:
            return render_template("index.html", trainee=None, error=str(e))

    return render_template("index.html", trainee=None, error=error)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

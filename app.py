from flask import Flask, render_template, request, redirect, url_for, send_file
import os, json, re, tempfile
import pandas as pd
from datetime import datetime

from docxtpl import DocxTemplate
from docx2pdf import convert  # على Render غالباً لن يعمل (ويندوز فقط) -> سنستخدم LibreOffice
import subprocess

app = Flask(__name__)

DATA_DIR = "data"
STUDENTS_PATH = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_PATH = os.path.join(DATA_DIR, "slots.json")
ENTITIES_PATH = os.path.join(DATA_DIR, "entities.json")
ASSIGNMENTS_PATH = os.path.join(DATA_DIR, "assignments.json")
TEMPLATE_DOCX_PATH = os.path.join(DATA_DIR, "letter_template.docx")


# ---------------- Utilities ----------------
def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except:
            return default

def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def normalize_phone(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = s.replace(" ", "").replace("-", "")
    # لو كان رقم أكسل منتهي .0
    if s.endswith(".0"):
        s = s[:-2]
    # استخراج أرقام فقط
    digits = "".join(ch for ch in s if ch.isdigit())
    return digits

def last4_ok(full_phone, last4):
    full_phone = normalize_phone(full_phone)
    last4 = str(last4).strip()
    return len(full_phone) >= 4 and full_phone[-4:] == last4

def read_students_df():
    if not os.path.exists(STUDENTS_PATH):
        raise FileNotFoundError(f"students.xlsx غير موجود: {STUDENTS_PATH}")
    df = pd.read_excel(STUDENTS_PATH)

    # أعمدة مطلوبة (عدّلها إذا أسماء أعمدتك مختلفة)
    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("الأعمدة الناقصة في students.xlsx: " + " | ".join(missing))

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()

    return df

def libreoffice_docx_to_pdf(docx_path, pdf_path):
    # تحويل عبر LibreOffice (يعمل على Linux/Render إذا كانت الحزمة موجودة)
    out_dir = os.path.dirname(pdf_path)
    os.makedirs(out_dir, exist_ok=True)

    cmd = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        "--outdir",
        out_dir,
        docx_path
    ]
    subprocess.check_call(cmd)

    # LibreOffice يسمي الملف بنفس اسم docx لكن pdf
    expected_pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if expected_pdf != pdf_path:
        # انقل الاسم النهائي المطلوب
        os.replace(expected_pdf, pdf_path)

# ---------------- Routes ----------------
@app.route("/", methods=["GET", "POST"])
def index():
    error = None

    # لتحميل الجهات فقط في صفحة الدخول (لو تحب تعرضها لاحقاً)
    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال"
            return render_template("index.html", error=error)

        try:
            df = read_students_df()
        except Exception as e:
            return render_template("index.html", error=str(e))

        row = df[df["رقم المتدرب"] == trainee_id]
        if row.empty:
            return render_template("index.html", error="لم يتم العثور على رقم المتدرب")

        trainee = row.iloc[0].to_dict()

        if not last4_ok(trainee.get("رقم الجوال_norm", ""), last4):
            return render_template("index.html", error="آخر 4 أرقام من الجوال غير صحيحة")

        # توجيه لصفحة الخيارات
        return redirect(url_for("dashboard", trainee_id=trainee_id, last4=last4))

    return render_template("index.html", error=error)

@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    trainee_id = (request.args.get("trainee_id") or "").strip()
    last4 = (request.args.get("last4") or "").strip()

    if not trainee_id or not last4:
        return redirect(url_for("index"))

    df = read_students_df()
    row = df[df["رقم المتدرب"] == trainee_id]
    if row.empty:
        return redirect(url_for("index"))

    trainee = row.iloc[0].to_dict()
    if not last4_ok(trainee.get("رقم الجوال_norm", ""), last4):
        return redirect(url_for("index"))

    specialty = trainee["التخصص"]

    entities = load_json(ENTITIES_PATH, [])
    slots = load_json(SLOTS_PATH, {})
    assignments = load_json(ASSIGNMENTS_PATH, {})

    # إذا المتدرب اختار سابقاً، نعرض اختياره وزر PDF
    existing = assignments.get(trainee_id)

    # المقاعد المتاحة لهذا التخصص فقط
    specialty_slots = slots.get(specialty, {})

    # فلترة الجهات: فقط الجهات التي لها مقاعد > 0 ضمن تخصص المتدرب
    available_entities = []
    for ent in entities:
        count = int(specialty_slots.get(ent, 0) or 0)
        if count > 0:
            available_entities.append({"name": ent, "count": count})

    error = None
    success = None

    if request.method == "POST":
        chosen_entity = (request.form.get("entity") or "").strip()

        if existing:
            error = "تم اختيار جهة مسبقاً لهذا المتدرب ولا يمكن التغيير."
        else:
            if not chosen_entity:
                error = "اختر جهة التدريب"
            else:
                current = int(specialty_slots.get(chosen_entity, 0) or 0)
                if current <= 0:
                    error = "لا توجد مقاعد متاحة لهذه الجهة ضمن تخصصك"
                else:
                    # نقص المقعد
                    specialty_slots[chosen_entity] = current - 1
                    slots[specialty] = specialty_slots
                    save_json(SLOTS_PATH, slots)

                    # سجل التعيين
                    assignments[trainee_id] = {
                        "trainee_id": trainee_id,
                        "trainee_name": trainee["اسم المتدرب"],
                        "specialty": specialty,
                        "phone": trainee.get("رقم الجوال_norm", ""),
                        "entity": chosen_entity,
                        "chosen_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    save_json(ASSIGNMENTS_PATH, assignments)

                    success = "تم حفظ اختيارك بنجاح ✅"
                    existing = assignments.get(trainee_id)

                    # تحديث القائمة بعد التغيير
                    specialty_slots = slots.get(specialty, {})
                    available_entities = []
                    for ent in entities:
                        count = int(specialty_slots.get(ent, 0) or 0)
                        if count > 0:
                            available_entities.append({"name": ent, "count": count})

    return render_template(
        "dashboard.html",
        trainee=trainee,
        specialty=specialty,
        existing=existing,
        available_entities=available_entities,
        error=error,
        success=success,
        trainee_id=trainee_id,
        last4=last4
    )

@app.route("/download_pdf/<trainee_id>")
def download_pdf(trainee_id):
    assignments = load_json(ASSIGNMENTS_PATH, {})
    info = assignments.get(trainee_id)
    if not info:
        return "لا يوجد تعيين لهذا المتدرب", 404

    if not os.path.exists(TEMPLATE_DOCX_PATH):
        return "ملف قالب الوورد غير موجود data/letter_template.docx", 500

    # توليد DOCX جديد من القالب
    tpl = DocxTemplate(TEMPLATE_DOCX_PATH)

    context = {
        "TraineeName": info.get("trainee_name", ""),
        "AcademicID": info.get("trainee_id", ""),
        "Phone": info.get("phone", ""),
        "Specialty": info.get("specialty", ""),
        "Company": info.get("entity", ""),
        "LetterNo": info.get("trainee_id", ""),  # لو عندك رقم خطاب حقيقي عدله
        "course_ref": "",
        "college_supervisor": ""
    }

    with tempfile.TemporaryDirectory() as tmp:
        out_docx = os.path.join(tmp, f"letter_{trainee_id}.docx")
        out_pdf = os.path.join(tmp, f"letter_{trainee_id}.pdf")

        tpl.render(context)
        tpl.save(out_docx)

        # تحويل PDF
        try:
            libreoffice_docx_to_pdf(out_docx, out_pdf)
        except Exception as e:
            return f"فشل تحويل PDF: {e}", 500

        return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")

if __name__ == "__main__":
    app.run(debug=True)

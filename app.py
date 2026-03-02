from flask import Flask, render_template, request, redirect, url_for, send_file
import os, json, tempfile, subprocess
import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate

app = Flask(__name__)

DATA_DIR = "data"
STUDENTS_PATH = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_PATH = os.path.join(DATA_DIR, "slots.json")
ENTITIES_PATH = os.path.join(DATA_DIR, "entities.json")
ASSIGNMENTS_PATH = os.path.join(DATA_DIR, "assignments.json")
TEMPLATE_DOCX_PATH = os.path.join(DATA_DIR, "letter_template.docx")


# ============ Helpers ============
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
    s = str(x).strip().replace(" ", "").replace("-", "")
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())

def last4_ok(full_phone, last4):
    p = normalize_phone(full_phone)
    last4 = str(last4).strip()
    return len(p) >= 4 and p[-4:] == last4

def read_students_df():
    if not os.path.exists(STUDENTS_PATH):
        raise FileNotFoundError(f"students.xlsx غير موجود: {STUDENTS_PATH}")

    df = pd.read_excel(STUDENTS_PATH)

    # عدّل أسماء الأعمدة هنا حسب ملفك لو اختلفت:
    COL_ID = "رقم المتدرب"
    COL_PHONE = "رقم الجوال"
    COL_NAME = "اسم المتدرب"
    COL_SPEC = "التخصص"
    COL_ENTITY = "جهة التدريب"   # <<< هذا أهم شيء

    required = [COL_ID, COL_PHONE, COL_NAME, COL_SPEC, COL_ENTITY]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("الأعمدة الناقصة في students.xlsx: " + " | ".join(missing))

    df[COL_ID] = df[COL_ID].astype(str).str.strip()
    df[COL_NAME] = df[COL_NAME].astype(str).str.strip()
    df[COL_SPEC] = df[COL_SPEC].astype(str).str.strip()
    df[COL_ENTITY] = df[COL_ENTITY].astype(str).str.strip()
    df["رقم_الجوال_norm"] = df[COL_PHONE].apply(normalize_phone)

    return df, COL_ID, COL_PHONE, COL_NAME, COL_SPEC, COL_ENTITY

def build_slots_and_entities_from_excel():
    """
    يبني:
      entities.json = قائمة الجهات
      slots.json = { تخصص: { جهة: عدد المقاعد } }
    من students.xlsx تلقائياً
    """
    df, COL_ID, COL_PHONE, COL_NAME, COL_SPEC, COL_ENTITY = read_students_df()

    # احذف الصفوف اللي جهة التدريب أو التخصص فاضية
    df = df[(df[COL_SPEC] != "") & (df[COL_ENTITY] != "")]

    # الجهات
    entities = sorted(df[COL_ENTITY].dropna().unique().tolist())
    save_json(ENTITIES_PATH, entities)

    # المقاعد = count لكل (تخصص + جهة)
    grouped = df.groupby([COL_SPEC, COL_ENTITY]).size().reset_index(name="count")

    slots = {}
    for _, row in grouped.iterrows():
        spec = str(row[COL_SPEC]).strip()
        ent = str(row[COL_ENTITY]).strip()
        cnt = int(row["count"])
        slots.setdefault(spec, {})[ent] = cnt

    save_json(SLOTS_PATH, slots)

    # assignments.json إذا غير موجود
    if not os.path.exists(ASSIGNMENTS_PATH):
        save_json(ASSIGNMENTS_PATH, {})

def ensure_auto_data_ready():
    """
    إذا slots/entities غير موجودة، يولدها من Excel تلقائياً
    """
    if (not os.path.exists(SLOTS_PATH)) or (not os.path.exists(ENTITIES_PATH)):
        build_slots_and_entities_from_excel()

def libreoffice_docx_to_pdf(docx_path, pdf_path):
    out_dir = os.path.dirname(pdf_path)
    os.makedirs(out_dir, exist_ok=True)
    cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path]
    subprocess.check_call(cmd)

    expected = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if expected != pdf_path:
        os.replace(expected, pdf_path)


# ============ Routes ============
@app.route("/", methods=["GET", "POST"])
def index():
    ensure_auto_data_ready()

    error = None
    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()
        if not trainee_id or not last4:
            return render_template("index.html", error="فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال")

        df, COL_ID, COL_PHONE, COL_NAME, COL_SPEC, COL_ENTITY = read_students_df()
        row = df[df[COL_ID] == trainee_id]
        if row.empty:
            return render_template("index.html", error="لم يتم العثور على رقم المتدرب")

        trainee = row.iloc[0].to_dict()
        if not last4_ok(trainee["رقم_الجوال_norm"], last4):
            return render_template("index.html", error="آخر 4 أرقام من الجوال غير صحيحة")

        return redirect(url_for("dashboard", trainee_id=trainee_id, last4=last4))

    return render_template("index.html", error=error)

@app.route("/dashboard", methods=["GET", "POST"])
def dashboard():
    ensure_auto_data_ready()

    trainee_id = (request.args.get("trainee_id") or "").strip()
    last4 = (request.args.get("last4") or "").strip()
    if not trainee_id or not last4:
        return redirect(url_for("index"))

    df, COL_ID, COL_PHONE, COL_NAME, COL_SPEC, COL_ENTITY = read_students_df()
    row = df[df[COL_ID] == trainee_id]
    if row.empty:
        return redirect(url_for("index"))

    trainee = row.iloc[0].to_dict()
    if not last4_ok(trainee["رقم_الجوال_norm"], last4):
        return redirect(url_for("index"))

    specialty = str(trainee[COL_SPEC]).strip()

    entities = load_json(ENTITIES_PATH, [])
    slots = load_json(SLOTS_PATH, {})
    assignments = load_json(ASSIGNMENTS_PATH, {})

    existing = assignments.get(trainee_id)

    specialty_slots = slots.get(specialty, {})
    available_entities = []
    for ent in entities:
        c = int(specialty_slots.get(ent, 0) or 0)
        if c > 0:
            available_entities.append({"name": ent, "count": c})

    error = None
    success = None

    if request.method == "POST":
        chosen_entity = (request.form.get("entity") or "").strip()

        if existing:
            error = "تم اختيار جهة مسبقاً ولا يمكن التغيير."
        else:
            if not chosen_entity:
                error = "اختر جهة التدريب"
            else:
                current = int(specialty_slots.get(chosen_entity, 0) or 0)
                if current <= 0:
                    error = "لا توجد مقاعد متاحة لهذه الجهة ضمن تخصصك"
                else:
                    # ينقص المقعد
                    specialty_slots[chosen_entity] = current - 1
                    slots[specialty] = specialty_slots
                    save_json(SLOTS_PATH, slots)

                    assignments[trainee_id] = {
                        "trainee_id": trainee_id,
                        "trainee_name": trainee[COL_NAME],
                        "specialty": specialty,
                        "phone": trainee["رقم_الجوال_norm"],
                        "entity": chosen_entity,
                        "chosen_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    save_json(ASSIGNMENTS_PATH, assignments)

                    success = "تم حفظ اختيارك بنجاح ✅"
                    existing = assignments.get(trainee_id)

                    # تحديث القائمة
                    specialty_slots = slots.get(specialty, {})
                    available_entities = []
                    for ent in entities:
                        c = int(specialty_slots.get(ent, 0) or 0)
                        if c > 0:
                            available_entities.append({"name": ent, "count": c})

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
    ensure_auto_data_ready()

    assignments = load_json(ASSIGNMENTS_PATH, {})
    info = assignments.get(trainee_id)
    if not info:
        return "لا يوجد تعيين لهذا المتدرب", 404

    if not os.path.exists(TEMPLATE_DOCX_PATH):
        return "ملف قالب الوورد غير موجود data/letter_template.docx", 500

    tpl = DocxTemplate(TEMPLATE_DOCX_PATH)
    context = {
        "TraineeName": info.get("trainee_name", ""),
        "AcademicID": info.get("trainee_id", ""),
        "Phone": info.get("phone", ""),
        "Specialty": info.get("specialty", ""),
        "Company": info.get("entity", ""),
        "LetterNo": info.get("trainee_id", ""),
        "course_ref": "",
        "college_supervisor": ""
    }

    with tempfile.TemporaryDirectory() as tmp:
        out_docx = os.path.join(tmp, f"letter_{trainee_id}.docx")
        out_pdf = os.path.join(tmp, f"letter_{trainee_id}.pdf")

        tpl.render(context)
        tpl.save(out_docx)

        try:
            libreoffice_docx_to_pdf(out_docx, out_pdf)
        except Exception as e:
            return f"فشل تحويل PDF: {e}", 500

        return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")


if __name__ == "__main__":
    app.run(debug=True)

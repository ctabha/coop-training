import os
import json
import tempfile
import subprocess
from datetime import datetime

from flask import Flask, render_template, request, redirect, url_for, send_file, abort
import pandas as pd
from docxtpl import DocxTemplate

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
TEMPLATE_DOCX = os.path.join(DATA_DIR, "letter_template.docx")

ENTITIES_JSON = os.path.join(DATA_DIR, "entities.json")
SLOTS_JSON = os.path.join(DATA_DIR, "slots.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")


# ----------------------------
# Helpers
# ----------------------------
def ensure_files_exist():
    missing = []
    for p in [STUDENTS_XLSX, TEMPLATE_DOCX]:
        if not os.path.exists(p):
            missing.append(p)
    if missing:
        raise FileNotFoundError("ملفات مفقودة: " + " , ".join(missing))

    # إذا ما فيه assignments.json ننشئه
    if not os.path.exists(ASSIGNMENTS_JSON):
        with open(ASSIGNMENTS_JSON, "w", encoding="utf-8") as f:
            json.dump({}, f, ensure_ascii=False, indent=2)

    # entities.json/slots.json اختيارية (سنولدها من الإكسل إذا غير موجودة)
    if not os.path.exists(ENTITIES_JSON):
        generate_entities_from_excel()

    if not os.path.exists(SLOTS_JSON):
        generate_slots_from_excel()


def read_students_df():
    df = pd.read_excel(STUDENTS_XLSX)

    # أسماء الأعمدة الأساسية المطلوبة (عدّلها إذا اسم العمود عندك مختلف)
    required = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "جهة التدريب"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"أعمدة ناقصة في Excel: {', '.join(missing)}")

    # تنظيف بسيط
    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال"] = df["رقم الجوال"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()

    return df


def normalize_phone_last4(phone: str) -> str:
    # نأخذ آخر 4 أرقام فعليًا من الجوال (حتى لو كان فيه + أو مسافات)
    digits = "".join(ch for ch in str(phone) if ch.isdigit())
    return digits[-4:] if len(digits) >= 4 else ""


def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_json(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


def generate_entities_from_excel():
    df = read_students_df()
    entities = sorted(set([x for x in df["جهة التدريب"].dropna().tolist() if str(x).strip() != ""]))
    data = {e: 0 for e in entities}  # نفس شكل ملفك (قيم 0)
    save_json(ENTITIES_JSON, data)


def generate_slots_from_excel():
    """
    المقاعد حسب التخصص = عدد المتدربين في كل تخصص (تقدر تغيّرها لاحقًا)
    """
    df = read_students_df()
    counts = df.groupby("التخصص").size().to_dict()
    slots = {str(k): int(v) for k, v in counts.items()}
    save_json(SLOTS_JSON, slots)


def get_remaining_slots(slots, assignments):
    """
    نحسب المتبقي لكل تخصص = slots[spec] - عدد المتدربين الذين تم تثبيت جهة لهم في هذا التخصص
    """
    used = {}
    for trainee_id, rec in assignments.items():
        spec = rec.get("specialty")
        entity = rec.get("entity")
        if spec and entity:
            used[spec] = used.get(spec, 0) + 1

    remaining = {}
    for spec, total in slots.items():
        remaining[spec] = int(total) - int(used.get(spec, 0))
    return remaining


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    تحويل DOCX إلى PDF باستخدام LibreOffice (يحتاج تثبيت libreoffice في render-build.sh)
    """
    cmd = [
        "libreoffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", os.path.dirname(pdf_path),
        docx_path,
    ]
    subprocess.check_call(cmd)
    # LibreOffice يخرج pdf بنفس اسم docx
    generated = os.path.join(os.path.dirname(pdf_path), os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    if not os.path.exists(generated):
        raise RuntimeError("لم يتم إنشاء PDF بواسطة LibreOffice")
    os.replace(generated, pdf_path)


# ----------------------------
# Routes
# ----------------------------
@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    ensure_files_exist()

    if request.method == "POST":
        trainee_id = (request.form.get("trainee_id") or "").strip()
        last4 = (request.form.get("last4") or "").strip()

        if not trainee_id or not last4:
            error = "فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال"
            return render_template("index.html", error=error)

        df = read_students_df()
        row = df[df["رقم المتدرب"] == trainee_id]
        if row.empty:
            error = "رقم المتدرب غير موجود"
            return render_template("index.html", error=error)

        trainee = row.iloc[0].to_dict()
        expected_last4 = normalize_phone_last4(trainee.get("رقم الجوال", ""))

        if expected_last4 != last4:
            error = "آخر 4 أرقام من الجوال غير صحيحة"
            return render_template("index.html", error=error)

        return redirect(url_for("dashboard", trainee_id=trainee_id))

    return render_template("index.html", error=error)


@app.route("/dashboard/<trainee_id>", methods=["GET", "POST"])
def dashboard(trainee_id):
    ensure_files_exist()

    df = read_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)]
    if row.empty:
        return "متدرب غير موجود", 404

    trainee = row.iloc[0].to_dict()
    specialty = str(trainee.get("التخصص", "")).strip()

    entities_dict = load_json(ENTITIES_JSON, {})
    slots = load_json(SLOTS_JSON, {})
    assignments = load_json(ASSIGNMENTS_JSON, {})

    # إعادة توليد تلقائي إذا صارت الملفات فاضية/غير منطقية
    if not entities_dict:
        generate_entities_from_excel()
        entities_dict = load_json(ENTITIES_JSON, {})
    if not slots:
        generate_slots_from_excel()
        slots = load_json(SLOTS_JSON, {})

    remaining = get_remaining_slots(slots, assignments)
    remaining_for_spec = remaining.get(specialty, 0)

    chosen = assignments.get(str(trainee_id))

    if request.method == "POST":
        entity = (request.form.get("entity") or "").strip()

        if remaining_for_spec <= 0:
            return render_template(
                "dashboard.html",
                trainee=trainee,
                entities=list(entities_dict.keys()),
                remaining_for_spec=remaining_for_spec,
                specialty=specialty,
                chosen=chosen,
                error="لا توجد مقاعد متبقية لهذا التخصص"
            )

        if not entity or entity not in entities_dict:
            return render_template(
                "dashboard.html",
                trainee=trainee,
                entities=list(entities_dict.keys()),
                remaining_for_spec=remaining_for_spec,
                specialty=specialty,
                chosen=chosen,
                error="اختر جهة صحيحة"
            )

        # إذا المتدرب اختار من قبل، لا ننقص مرتين
        if chosen and chosen.get("entity"):
            # تحديث الجهة فقط بدون إنقاص إضافي؟ (هنا نمنع العبث)
            return render_template(
                "dashboard.html",
                trainee=trainee,
                entities=list(entities_dict.keys()),
                remaining_for_spec=remaining_for_spec,
                specialty=specialty,
                chosen=chosen,
                error="تم تثبيت اختيارك مسبقًا. إذا تريد تعديل النظام للسماح بالتعديل قل لي."
            )

        assignments[str(trainee_id)] = {
            "trainee_id": str(trainee_id),
            "name": trainee.get("اسم المتدرب", ""),
            "specialty": specialty,
            "entity": entity,
            "timestamp": datetime.utcnow().isoformat(),
        }
        save_json(ASSIGNMENTS_JSON, assignments)

        # بعد الحفظ نعيد حساب المتبقي
        remaining = get_remaining_slots(slots, assignments)
        remaining_for_spec = remaining.get(specialty, 0)
        chosen = assignments.get(str(trainee_id))

    return render_template(
        "dashboard.html",
        trainee=trainee,
        entities=list(entities_dict.keys()),
        remaining_for_spec=remaining_for_spec,
        specialty=specialty,
        chosen=chosen,
        error=None
    )


@app.route("/download/<trainee_id>", methods=["GET"])
def download(trainee_id):
    ensure_files_exist()

    assignments = load_json(ASSIGNMENTS_JSON, {})
    rec = assignments.get(str(trainee_id))
    if not rec or not rec.get("entity"):
        return "لا يوجد اختيار جهة لهذا المتدرب", 400

    df = read_students_df()
    row = df[df["رقم المتدرب"] == str(trainee_id)]
    if row.empty:
        return "متدرب غير موجود", 404
    trainee = row.iloc[0].to_dict()

    # تعبئة قالب الوورد (عدّل أسماء الحقول لتطابق placeholders في letter_template.docx)
    tpl = DocxTemplate(TEMPLATE_DOCX)
    context = {
        "TraineeName": trainee.get("اسم المتدرب", ""),
        "AcademicID": trainee.get("رقم المتدرب", ""),
        "Specialty": trainee.get("التخصص", ""),
        "Phone": trainee.get("رقم الجوال", ""),
        "Company": rec.get("entity", ""),
        "LetterNo": trainee.get("رقم المتدرب", ""),  # مثال: تقدر تغيّره
    }
    tpl.render(context)

    with tempfile.TemporaryDirectory() as tmp:
        out_docx = os.path.join(tmp, f"letter_{trainee_id}.docx")
        out_pdf = os.path.join(tmp, f"letter_{trainee_id}.pdf")
        tpl.save(out_docx)

        # نحاول PDF (إذا LibreOffice موجود)
        try:
            convert_docx_to_pdf(out_docx, out_pdf)
            return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.pdf")
        except Exception:
            # fallback: نرجع DOCX بدل ما نخلي المستخدم بدون ملف
            return send_file(out_docx, as_attachment=True, download_name=f"خطاب_توجيه_{trainee_id}.docx")


@app.route("/health")
def health():
    return "ok", 200


if __name__ == "__main__":
    app.run(debug=True)

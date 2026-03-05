import os
import io
import json
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash

import pandas as pd
from docx import Document

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ENTITIES_JSON = os.path.join(DATA_DIR, "entities.json")
SLOTS_JSON = os.path.join(DATA_DIR, "slots.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")
LETTER_TEMPLATE_DOCX = os.path.join(DATA_DIR, "letter_template.docx")


# ---------------------------
# Helpers: Safe JSON read/write
# ---------------------------
def safe_load_json(path, default):
    try:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        # لو الملف فاسد/فيه فاصلة/اقتباسات غلط... لا نطيّح الموقع
        return default


def safe_write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ---------------------------
# Load Students (Fix KeyError: 0)
# ---------------------------
def load_students():
    # لازم الأعمدة تكون موجودة بالاسم العربي اللي عندك
    # من صور إكسل عندك، الأعمدة مثل:
    # (رقم المتدرب) (اسم المتدرب) (رقم الجوال) (التخصص) (البرنامج) (المدرب) ... الخ
    df = pd.read_excel(STUDENTS_XLSX)

    # تنظيف أسماء الأعمدة من الفراغات
    df.columns = [str(c).strip() for c in df.columns]

    # حاول نلتقط أسماء الأعمدة الأكثر شيوعًا عندك:
    col_id = None
    col_name = None
    col_phone = None
    col_spec = None
    col_program = None
    col_city = None
    col_ref = None

    # جرّب أسماء متوقعة
    for c in df.columns:
        if col_id is None and ("رقم المتدرب" in c or "الرقم التدريبي" in c or c.strip() == "رقم المتدرب"):
            col_id = c
        if col_name is None and ("اسم المتدرب" in c or c.strip() == "اسم المتدرب"):
            col_name = c
        if col_phone is None and ("رقم الجوال" in c or "الجوال" in c or c.strip() == "رقم الجوال"):
            col_phone = c
        if col_spec is None and ("التخصص" in c or c.strip() == "التخصص"):
            col_spec = c
        if col_program is None and ("البرنامج" in c or c.strip() == "برنامج"):
            col_program = c
        if col_city is None and ("المدينة" in c or c.strip() == "المدينة"):
            col_city = c
        if col_ref is None and ("الرقم المرجعي" in c or c.strip() == "الرقم المرجعي"):
            col_ref = c

    # لو ما لقى عمود معيّن، نخليه فاضي بدل ما يطيح
    students = []
    for _, r in df.iterrows():
        trainee_id = str(r[col_id]).strip() if col_id and pd.notna(r.get(col_id)) else ""
        trainee_name = str(r[col_name]).strip() if col_name and pd.notna(r.get(col_name)) else ""
        phone = str(r[col_phone]).strip() if col_phone and pd.notna(r.get(col_phone)) else ""
        spec = str(r[col_spec]).strip() if col_spec and pd.notna(r.get(col_spec)) else ""
        program = str(r[col_program]).strip() if col_program and pd.notna(r.get(col_program)) else ""
        city = str(r[col_city]).strip() if col_city and pd.notna(r.get(col_city)) else ""
        course_ref = str(r[col_ref]).strip() if col_ref and pd.notna(r.get(col_ref)) else ""

        # تجاهل الصفوف الفاضية
        if not trainee_id:
            continue

        students.append({
            "trainee_id": trainee_id,
            "trainee_name": trainee_name,
            "phone": phone,
            "spec": spec,
            "program": program,
            "city": city,
            "course_ref": course_ref,
        })

    return students


def find_student(trainee_id):
    students = load_students()
    for s in students:
        if str(s["trainee_id"]).strip() == str(trainee_id).strip():
            return s
    return None


# ---------------------------
# Entities / Slots / Assignments
# ---------------------------
def load_entities():
    # entities.json شكلها المتوقع:
    # {
    #   "تقنية المحركات ومركبات": [
    #       {"name": "شركة X", "college_supervisor": "....", "start_date": "...", "end_date": "..."},
    #       ...
    #   ],
    #   "تخصص آخر": [...]
    # }
    return safe_load_json(ENTITIES_JSON, {})


def load_slots():
    # slots.json شكلها المتوقع:
    # {"تقنية المحركات ومركبات": {"شركة X": 3, "شركة Y": 1}, "تخصص آخر": {...}}
    return safe_load_json(SLOTS_JSON, {})


def load_assignments():
    # assignments.json شكلها المتوقع:
    # {"444242291": {"spec": "...", "entity": "شركة X", "ts": "..."}, ...}
    # لو الملف فاسد/فاضي غلط -> يرجع {}
    return safe_load_json(ASSIGNMENTS_JSON, {})


def save_assignments(assignments):
    safe_write_json(ASSIGNMENTS_JSON, assignments)


def get_remaining_for_student(student, slots, assignments):
    """
    يرجع قائمة جهات تدريب لتخصص الطالب + عدد الفرص المتبقية لكل جهة.
    """
    spec = student.get("spec", "").strip()
    spec_slots = slots.get(spec, {})
    if not isinstance(spec_slots, dict):
        spec_slots = {}

    # احسب المستخدم لكل جهة لنفس التخصص
    used_by_entity = {}
    for tid, info in assignments.items():
        if not isinstance(info, dict):
            continue
        if info.get("spec") != spec:
            continue
        ent = info.get("entity")
        if not ent:
            continue
        used_by_entity[ent] = used_by_entity.get(ent, 0) + 1

    remaining = []
    for ent_name, total in spec_slots.items():
        try:
            total_int = int(total)
        except Exception:
            total_int = 0
        used_int = int(used_by_entity.get(ent_name, 0))
        rem = max(total_int - used_int, 0)
        remaining.append({"entity": ent_name, "total": total_int, "used": used_int, "remaining": rem})

    # رتب من الأعلى للأقل
    remaining.sort(key=lambda x: x["remaining"], reverse=True)
    return remaining


def pick_auto_entity(remaining_list):
    """
    اختيار تلقائي: أول جهة فيها remaining > 0
    """
    for item in remaining_list:
        if item["remaining"] > 0:
            return item["entity"]
    return ""


# ---------------------------
# DOCX Template Fill
# ---------------------------
def replace_in_paragraph(paragraph, mapping):
    # استبدال داخل paragraph مع الحفاظ قدر الإمكان على التنسيق
    # لأن الـ placeholders قد يكونون موزعين على Runs
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    new_text = full_text
    for k, v in mapping.items():
        new_text = new_text.replace(k, v)

    if new_text != full_text:
        # امسح الرنز ثم اكتب نص واحد
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = new_text if paragraph.runs else paragraph.add_run(new_text).text


def replace_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_in_paragraph(p, mapping)
            for t in cell.tables:
                replace_in_table(t, mapping)


def fill_letter_docx(student, chosen_entity, entities_data):
    doc = Document(LETTER_TEMPLATE_DOCX)

    # احضر بيانات الجهة من entities.json إن وجدت
    spec = student.get("spec", "").strip()
    ent_info = {}
    spec_entities = entities_data.get(spec, [])
    if isinstance(spec_entities, list):
        for e in spec_entities:
            if isinstance(e, dict) and e.get("name") == chosen_entity:
                ent_info = e
                break

    mapping = {
        "{{phone}}": str(student.get("phone", "")).strip(),
        "{{trainee_name}}": str(student.get("trainee_name", "")).strip(),
        "{{trainee_id}}": str(student.get("trainee_id", "")).strip(),
        "{{course_ref}}": str(student.get("course_ref", "")).strip(),
        "{{college_supervisor}}": str(ent_info.get("college_supervisor", "")).strip(),
        "{{training_entity}}": str(chosen_entity).strip(),
        "{{start_date}}": str(ent_info.get("start_date", "")).strip(),
        "{{end_date}}": str(ent_info.get("end_date", "")).strip(),
    }

    # Paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)

    # Tables
    for t in doc.tables:
        replace_in_table(t, mapping)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


# ---------------------------
# Routes
# ---------------------------
@app.get("/")
def index():
    return render_template("index.html")


@app.post("/")
def login():
    trainee_id = request.form.get("trainee_id", "").strip()
    last4 = request.form.get("last4", "").strip()

    student = find_student(trainee_id)
    if not student:
        flash("الرقم التدريبي غير موجود.")
        return redirect(url_for("index"))

    phone = str(student.get("phone", "")).strip()
    if len(phone) < 4 or phone[-4:] != last4:
        flash("آخر 4 أرقام من الجوال غير صحيحة.")
        return redirect(url_for("index"))

    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.get("/dashboard/<trainee_id>")
def dashboard(trainee_id):
    student = find_student(trainee_id)
    if not student:
        return "Student not found", 404

    slots = load_slots()
    assignments = load_assignments()
    entities = load_entities()

    remaining_list = get_remaining_for_student(student, slots, assignments)

    # إذا الطالب مسجّل سابقًا
    existing = assignments.get(str(trainee_id))
    existing_entity = ""
    if isinstance(existing, dict):
        existing_entity = existing.get("entity", "") or ""

    total_remaining = sum(x["remaining"] for x in remaining_list)

    return render_template(
        "dashboard.html",
        student=student,
        remaining_list=remaining_list,
        total_remaining=total_remaining,
        existing_entity=existing_entity
    )


@app.post("/choose/<trainee_id>")
def choose(trainee_id):
    student = find_student(trainee_id)
    if not student:
        return "Student not found", 404

    slots = load_slots()
    assignments = load_assignments()

    remaining_list = get_remaining_for_student(student, slots, assignments)

    chosen_entity = request.form.get("training_entity", "").strip()
    auto_pick = request.form.get("auto_pick", "").strip()

    if auto_pick == "1":
        chosen_entity = pick_auto_entity(remaining_list)

    if not chosen_entity:
        flash("اختر جهة تدريب أو استخدم الاختيار التلقائي.")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # تحقق إن فيها فرصة
    rem_map = {x["entity"]: x["remaining"] for x in remaining_list}
    if rem_map.get(chosen_entity, 0) <= 0:
        flash("هذه الجهة مكتملة ولا يوجد فرص متبقية. اختر جهة أخرى.")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # سجّل اختيار الطالب
    assignments[str(trainee_id)] = {
        "spec": student.get("spec", "").strip(),
        "entity": chosen_entity,
        "ts": datetime.utcnow().isoformat()
    }
    save_assignments(assignments)

    flash("تم تسجيل اختيارك بنجاح ✅")
    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.get("/download-letter/<trainee_id>")
def download_letter(trainee_id):
    student = find_student(trainee_id)
    if not student:
        return "Student not found", 404

    assignments = load_assignments()
    info = assignments.get(str(trainee_id))
    if not isinstance(info, dict) or not info.get("entity"):
        flash("لا يمكن طباعة الخطاب قبل اختيار جهة تدريب.")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    chosen_entity = info["entity"]
    entities = load_entities()

    out_docx = fill_letter_docx(student, chosen_entity, entities)

    filename = f"خطاب_توجيه_{trainee_id}.docx"
    return send_file(
        out_docx,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


# Health check
@app.get("/health")
def health():
    return {"ok": True}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "5000")))

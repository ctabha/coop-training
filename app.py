import os
import json
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import pandas as pd
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
TEMPL_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
ENTITIES_JSON = os.path.join(DATA_DIR, "entities.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")
LETTER_TEMPLATE = os.path.join(DATA_DIR, "letter_template.docx")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-secret-key-change-me")


# ---------------------------
# Helpers: Safe JSON load/save
# ---------------------------
def _safe_load_json(path, default):
    try:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            txt = f.read().strip()
            if not txt:
                return default
            return json.loads(txt)
    except json.JSONDecodeError:
        # لو الملف صار فيه كسر بالـ JSON
        return default
    except Exception:
        return default


def _safe_save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


# ---------------------------
# Loaders
# ---------------------------
def load_entities():
    # entities.json لازم يكون List
    data = _safe_load_json(ENTITIES_JSON, [])
    if isinstance(data, list):
        return data
    return []


def load_assignments():
    # assignments.json لازم يكون Dict {trainee_id: {...}}
    data = _safe_load_json(ASSIGNMENTS_JSON, {})
    if isinstance(data, dict):
        return data
    return {}


def load_students():
    """
    يقرأ students.xlsx بشكل مرن سواء أعمدة عربية أو إنجليزية.
    يرجع List of dict:
      trainee_id, trainee_name, phone, specialization, program, course_ref, college_supervisor
    """
    if not os.path.exists(STUDENTS_XLSX):
        return []

    df = pd.read_excel(STUDENTS_XLSX)

    # تنظيف أسماء الأعمدة
    df.columns = [str(c).strip() for c in df.columns]

    # خرائط محتملة للأعمدة (عربي/إنجليزي)
    candidates = {
        "trainee_id": ["trainee_id", "Trainee ID", "رقم المتدرب", "الرقم التدريبي", "رقم التدريب", "رقم"],
        "trainee_name": ["trainee_name", "Trainee Name", "اسم المتدرب", "اسم", "الاسم"],
        "phone": ["phone", "Phone", "رقم الجوال", "الجوال", "الهاتف", "رقم الهاتف"],
        "specialization": ["specialization", "Specialization", "التخصص"],
        "program": ["program", "Program", "البرنامج"],
        "course_ref": ["course_ref", "Course Ref", "الرقم المرجعي للمقرر", "مرجع المقرر", "رقم المقرر"],
        "college_supervisor": ["college_supervisor", "College Supervisor", "مشرف الكلية", "المشرف", "المرشد"],
    }

    def find_col(keys):
        for k in keys:
            if k in df.columns:
                return k
        return None

    col_map = {field: find_col(keys) for field, keys in candidates.items()}

    # لو ما حصل trainee_id / trainee_name لازم نوقف
    if not col_map["trainee_id"] or not col_map["trainee_name"]:
        return []

    students = []
    for _, row in df.iterrows():
        tid = str(row[col_map["trainee_id"]]).strip() if col_map["trainee_id"] else ""
        name = str(row[col_map["trainee_name"]]).strip() if col_map["trainee_name"] else ""
        if tid == "" or tid.lower() == "nan":
            continue
        if name == "" or name.lower() == "nan":
            name = "—"

        def get_val(field):
            c = col_map.get(field)
            if not c:
                return ""
            v = row[c]
            if pd.isna(v):
                return ""
            return str(v).strip()

        students.append({
            "trainee_id": tid,
            "trainee_name": name,
            "phone": get_val("phone"),
            "specialization": get_val("specialization"),
            "program": get_val("program"),
            "course_ref": get_val("course_ref"),
            "college_supervisor": get_val("college_supervisor"),
        })

    return students


# ---------------------------
# Slots calculation (AUTO from entities + assignments)
# ---------------------------
def compute_remaining_from_entities(entities, assignments, students):
    """
    entities: list of {name, specialization, slots}
    assignments: dict {trainee_id: {entity, ...}}
    students: list of student dict
    returns:
      remaining_by_spec = {spec: {entity_name: remaining_int}}
      total_by_spec     = {spec: total_slots_int}
    """
    total_by_spec = {}
    total_by_entity = {}  # (spec, entity) -> total

    for e in entities:
        spec = str(e.get("specialization", "")).strip()
        name = str(e.get("name", "")).strip()
        if not spec or not name:
            continue
        try:
            slots = int(e.get("slots", 0) or 0)
        except Exception:
            slots = 0

        total_by_spec[spec] = total_by_spec.get(spec, 0) + slots
        total_by_entity[(spec, name)] = slots

    id_to_student = {s["trainee_id"]: s for s in students}
    used_by_entity = {}

    for tid, a in (assignments or {}).items():
        a = a or {}
        ent = a.get("entity")
        if not ent:
            continue
        st = id_to_student.get(str(tid))
        if not st:
            continue
        spec = str(st.get("specialization", "")).strip()
        if not spec:
            continue
        key = (spec, str(ent).strip())
        used_by_entity[key] = used_by_entity.get(key, 0) + 1

    remaining_by_spec = {}
    for (spec, name), total in total_by_entity.items():
        used = used_by_entity.get((spec, name), 0)
        rem = total - used
        if rem < 0:
            rem = 0
        remaining_by_spec.setdefault(spec, {})
        remaining_by_spec[spec][name] = rem

    return remaining_by_spec, total_by_spec


def available_entities_for_student(student, entities, assignments, students_all):
    """
    يرجّع قائمة الجهات المتاحة (remaining > 0) لنفس تخصص الطالب.
    """
    spec = str(student.get("specialization", "")).strip()
    remaining_by_spec, _ = compute_remaining_from_entities(entities, assignments, students_all)
    remaining = remaining_by_spec.get(spec, {})

    # قائمة الجهات من entities بنفس التخصص + فيها remaining > 0
    result = []
    for e in entities:
        if str(e.get("specialization", "")).strip() != spec:
            continue
        name = str(e.get("name", "")).strip()
        if not name:
            continue
        if remaining.get(name, 0) > 0:
            result.append({
                "name": name,
                "remaining": remaining.get(name, 0),
                "slots": int(e.get("slots", 0) or 0),
            })

    # ترتيب: الأكبر فرص أولاً
    result.sort(key=lambda x: x["remaining"], reverse=True)
    return result


# ---------------------------
# DOCX template filling
# ---------------------------
def _replace_in_paragraph(paragraph, mapping):
    # استبدال نص بسيط (يدعم وجود النص داخل run واحد غالبًا)
    for key, val in mapping.items():
        if key in paragraph.text:
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, val)


def _replace_everywhere(doc: Document, mapping):
    # body
    for p in doc.paragraphs:
        _replace_in_paragraph(p, mapping)

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, mapping)

    # headers/footers
    for section in doc.sections:
        header = section.header
        footer = section.footer
        for p in header.paragraphs:
            _replace_in_paragraph(p, mapping)
        for p in footer.paragraphs:
            _replace_in_paragraph(p, mapping)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_in_paragraph(p, mapping)
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        _replace_in_paragraph(p, mapping)


def generate_letter_docx(student, entity_name):
    """
    يولّد ملف Word على نفس قالبك letter_template.docx
    """
    if not os.path.exists(LETTER_TEMPLATE):
        raise FileNotFoundError("letter_template.docx غير موجود داخل data/")

    doc = Document(LETTER_TEMPLATE)

    # قيم افتراضية للتواريخ (تقدر تغيرها حسب ملفك لو عندك start_date/end_date في students.xlsx)
    start_date = student.get("start_date", "") or ""
    end_date = student.get("end_date", "") or ""

    mapping = {
        "{{trainee_name}}": student.get("trainee_name", ""),
        "{{trainee_id}}": student.get("trainee_id", ""),
        "{{phone}}": student.get("phone", ""),
        "{{specialization}}": student.get("specialization", ""),
        "{{program}}": student.get("program", ""),
        "{{course_ref}}": student.get("course_ref", ""),
        "{{college_supervisor}}": student.get("college_supervisor", ""),
        "{{training_entity}}": entity_name,
        "{{start_date}}": start_date,
        "{{end_date}}": end_date,
    }

    _replace_everywhere(doc, mapping)

    out_name = f"letter_{student.get('trainee_id','')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
    out_path = os.path.join(OUTPUT_DIR, out_name)
    doc.save(out_path)
    return out_path


# ---------------------------
# Routes
# ---------------------------
@app.route("/", methods=["GET", "POST", "HEAD"])
def index():
    # Render health check sends HEAD sometimes
    if request.method == "HEAD":
        return ("", 200)

    students = load_students()

    if request.method == "GET":
        return render_template("index.html")

    trainee_id = (request.form.get("trainee_id") or "").strip()
    phone_last4 = (request.form.get("phone_last4") or "").strip()

    if not trainee_id:
        flash("اكتب رقم المتدرب.", "error")
        return redirect(url_for("index"))

    # البحث عن الطالب
    st = None
    for s in students:
        if s["trainee_id"] == trainee_id:
            st = s
            break

    if not st:
        flash("رقم المتدرب غير موجود في ملف الطلاب.", "error")
        return redirect(url_for("index"))

    # لو رقم الجوال موجود، نتحقق من آخر 4
    phone = (st.get("phone") or "").strip()
    if phone_last4:
        if len(phone) < 4 or not phone.endswith(phone_last4):
            flash("آخر 4 أرقام من الجوال غير صحيحة.", "error")
            return redirect(url_for("index"))

    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    students_all = load_students()
    assignments = load_assignments()
    entities = load_entities()

    student = next((s for s in students_all if s["trainee_id"] == str(trainee_id)), None)
    if not student:
        return "Trainee not found", 404

    spec = str(student.get("specialization", "")).strip()

    remaining_by_spec, total_by_spec = compute_remaining_from_entities(entities, assignments, students_all)
    total_slots = total_by_spec.get(spec, 0)

    # الجهات المتاحة (المتبقي > 0)
    available = available_entities_for_student(student, entities, assignments, students_all)

    # التعيين الحالي لو موجود
    assigned = assignments.get(str(trainee_id))

    return render_template(
        "dashboard.html",
        student=student,
        total_slots=total_slots,
        available=available,
        assigned=assigned
    )


@app.route("/assign/<trainee_id>", methods=["POST"])
def assign(trainee_id):
    students_all = load_students()
    assignments = load_assignments()
    entities = load_entities()

    student = next((s for s in students_all if s["trainee_id"] == str(trainee_id)), None)
    if not student:
        flash("المتدرب غير موجود.", "error")
        return redirect(url_for("index"))

    chosen = (request.form.get("entity") or "").strip()
    if not chosen:
        flash("اختر جهة تدريب.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # تحقق أن الجهة متاحة فعلاً (باقي فيها فرص)
    available = available_entities_for_student(student, entities, assignments, students_all)
    allowed_names = {a["name"] for a in available}
    if chosen not in allowed_names:
        flash("هذه الجهة لا توجد بها فرص متبقية.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    # حفظ التعيين
    assignments[str(trainee_id)] = {
        "entity": chosen,
        "assigned_at": datetime.now().isoformat(timespec="seconds")
    }
    _safe_save_json(ASSIGNMENTS_JSON, assignments)

    flash("تم تأكيد الاختيار وتنقيص الفرص تلقائيًا.", "ok")
    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.route("/auto_assign/<trainee_id>", methods=["POST"])
def auto_assign(trainee_id):
    students_all = load_students()
    assignments = load_assignments()
    entities = load_entities()

    student = next((s for s in students_all if s["trainee_id"] == str(trainee_id)), None)
    if not student:
        flash("المتدرب غير موجود.", "error")
        return redirect(url_for("index"))

    available = available_entities_for_student(student, entities, assignments, students_all)
    if not available:
        flash("لا توجد جهات متاحة في تخصصك.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    chosen = available[0]["name"]  # أول جهة فيها فرص
    assignments[str(trainee_id)] = {
        "entity": chosen,
        "assigned_at": datetime.now().isoformat(timespec="seconds")
    }
    _safe_save_json(ASSIGNMENTS_JSON, assignments)

    flash(f"تم اختيار جهة تلقائيًا: {chosen}", "ok")
    return redirect(url_for("dashboard", trainee_id=trainee_id))


@app.route("/download_letter/<trainee_id>", methods=["GET"])
def download_letter(trainee_id):
    students_all = load_students()
    assignments = load_assignments()

    student = next((s for s in students_all if s["trainee_id"] == str(trainee_id)), None)
    if not student:
        return "Trainee not found", 404

    assigned = assignments.get(str(trainee_id))
    if not assigned or not assigned.get("entity"):
        flash("لا يوجد اختيار محفوظ، اختر جهة أولاً.", "error")
        return redirect(url_for("dashboard", trainee_id=trainee_id))

    entity_name = assigned["entity"]

    try:
        out_path = generate_letter_docx(student, entity_name)
    except Exception as e:
        return f"Error generating letter: {e}", 500

    return send_file(out_path, as_attachment=True, download_name="خطاب_التوجيه.docx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

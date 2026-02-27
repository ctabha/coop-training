import os
import json
import pandas as pd
from flask import Flask, request, send_file, jsonify, render_template
from docxtpl import DocxTemplate
from io import BytesIO

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")

STUDENTS_FILE = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_FILE = os.path.join(DATA_DIR, "slots.json")
ASSIGNMENTS_FILE = os.path.join(DATA_DIR, "assignments.json")
TEMPLATE_FILE = os.path.join(DATA_DIR, "letter_template.docx")


def load_students():
    df = pd.read_excel(STUDENTS_FILE)
    # تنظيف أسماء الأعمدة (اختياري لكن مفيد)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_json(path, default):
    if not os.path.exists(path):
        return default
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


@app.get("/api/slots")
def api_slots():
    # ترجع الفرص للواجهة
    slots = load_json(SLOTS_FILE, [])
    return jsonify(slots)


@app.post("/api/assign")
def api_assign():
    """
    ينقص الفرصة/يحجزها للمتدرب
    body: { "trainee_id": "...", "slot_id": "..." }
    """
    body = request.get_json(force=True)
    trainee_id = str(body.get("trainee_id", "")).strip()
    slot_id = str(body.get("slot_id", "")).strip()

    assignments = load_json(ASSIGNMENTS_FILE, {})
    slots = load_json(SLOTS_FILE, [])

    if not trainee_id or not slot_id:
        return jsonify({"ok": False, "error": "بيانات ناقصة"}), 400

    # تحقق من وجود الفرصة
    slot = next((s for s in slots if str(s.get("id")) == slot_id), None)
    if not slot:
        return jsonify({"ok": False, "error": "الفرصة غير موجودة"}), 404

    # تحقق من المتاح
    remaining = int(slot.get("remaining", 0))
    if remaining <= 0:
        return jsonify({"ok": False, "error": "لا يوجد مقاعد متاحة"}), 400

    # منع الحجز المكرر
    if trainee_id in assignments:
        return jsonify({"ok": False, "error": "تم الحجز سابقًا"}), 400

    # احجز وانقص
    assignments[trainee_id] = {"slot_id": slot_id}
    slot["remaining"] = remaining - 1

    # احفظ
    with open(ASSIGNMENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(assignments, f, ensure_ascii=False, indent=2)

    with open(SLOTS_FILE, "w", encoding="utf-8") as f:
        json.dump(slots, f, ensure_ascii=False, indent=2)

    return jsonify({"ok": True, "remaining": slot["remaining"]})


@app.post("/letter")
def generate_letter():
    """
    يطلع خطاب DOCX معبّى
    form fields: trainee_id, phone_last4
    """
    trainee_id = str(request.form.get("trainee_id", "")).strip()
    phone_last4 = str(request.form.get("phone_last4", "")).strip()

    if not trainee_id or not phone_last4:
        return "البيانات ناقصة", 400

    if not os.path.exists(TEMPLATE_FILE):
        return "قالب الخطاب غير موجود", 500

    df = load_students()

    # عدّل أسماء الأعمدة هنا حسب ملفك (من صورتك)
    # عندك: رقم المتدرب, رقم الجوال, اسم المتدرب, الرقم المرجعي, اسم المقرر, المدرب, جهة التدريب
    id_col = "رقم المتدرب"
    phone_col = "رقم الجوال"
    name_col = "اسم المتدرب"
    ref_col = "الرقم المرجعي"
    course_col = "اسم المقرر"
    supervisor_col = "المدرب"
    entity_col = "جهة التدريب"

    for col in [id_col, phone_col, name_col]:
        if col not in df.columns:
            return f"العمود غير موجود في students.xlsx: {col}", 500

    # طابق المتدرب
    df[id_col] = df[id_col].astype(str).str.strip()
    df[phone_col] = df[phone_col].astype(str).str.strip()

    row = df[df[id_col] == trainee_id]
    if row.empty:
        return "لم يتم العثور على متدرب بهذا الرقم", 404

    row = row.iloc[0]
    phone = str(row.get(phone_col, "")).strip()
    if len(phone) < 4 or phone[-4:] != phone_last4:
        return "آخر 4 أرقام من الجوال غير صحيحة", 400

    # ✅ المفاتيح هنا لازم تطابق اللي داخل قالب الـ DOCX عندك
    context = {
        "phone": phone,
        "trainee_name": str(row.get(name_col, "")).strip(),
        "trainee_id": trainee_id,
        "course_ref": str(row.get(ref_col, "")).strip(),
        "college_supervisor": str(row.get(supervisor_col, "")).strip(),
        "training_entity": str(row.get(entity_col, "")).strip(),
    }

    doc = DocxTemplate(TEMPLATE_FILE)
    doc.render(context)

    out = BytesIO()
    doc.save(out)
    out.seek(0)

    filename = f"letter_{trainee_id}.docx"
    return send_file(out, as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

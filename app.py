from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import os
import json
import tempfile
import subprocess
from datetime import datetime

app = Flask(__name__)

DATA_XLSX = os.path.join("data", "students.xlsx")
TEMPLATE_DOCX = os.path.join("data", "letter_template.docx")
SLOTS_JSON = os.path.join("data", "slots.json")
ASSIGNMENTS_JSON = os.path.join("data", "assignments.json")

REQUIRED_COLS = ["رقم المتدرب", "رقم الجوال", "اسم المتدرب", "التخصص", "البرنامج", "جهة التدريب"]

def load_json(path, default):
    try:
        if not os.path.exists(path):
            return default
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def save_json(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def normalize_phone(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    # إزالة .0 لو كانت أرقام من إكسل
    if s.endswith(".0"):
        s = s[:-2]
    # خذ فقط الأرقام
    s = "".join(ch for ch in s if ch.isdigit())
    return s

def read_students():
    if not os.path.exists(DATA_XLSX):
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {DATA_XLSX}")

    df = pd.read_excel(DATA_XLSX)
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError("الأعمدة التالية غير موجودة في Excel: " + " | ".join(missing))

    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال_norm"] = df["رقم الجوال"].apply(normalize_phone)
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).fillna("").str.strip()
    df["التخصص"] = df["التخصص"].astype(str).fillna("").str.strip()
    df["البرنامج"] = df["البرنامج"].astype(str).fillna("").str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).fillna("").str.strip()
    return df

def build_slots_from_excel(df: pd.DataFrame):
    """
    كل صف في Excel يعتبر فرصة:
    (التخصص + جهة التدريب) = مقعد واحد.
    نجمعها ونطلع عدد المقاعد لكل جهة داخل كل تخصص.
    """
    slots = {}
    grouped = df[df["جهة التدريب"] != ""].groupby(["التخصص", "جهة التدريب"]).size()
    for (spec, ent), count in grouped.items():
        spec = str(spec).strip()
        ent = str(ent).strip()
        slots.setdefault(spec, {})
        slots[spec][ent] = int(count)
    return slots

def ensure_slots(df: pd.DataFrame):
    slots = load_json(SLOTS_JSON, {})
    if not isinstance(slots, dict) or not slots:
        slots = build_slots_from_excel(df)
        save_json(SLOTS_JSON, slots)
    return slots

def total_and_remaining_for_specialty(slots: dict, specialty: str):
    spec_slots = slots.get(specialty, {})
    total = sum(int(v) for v in spec_slots.values())
    remaining = sum(int(v) for v in spec_slots.values())
    return total, remaining

def get_entities_for_specialty(slots: dict, specialty: str):
    spec_slots = slots.get(specialty, {})
    items = []
    for ent, cnt in spec_slots.items():
        cnt = int(cnt)
        if cnt > 0:
            items.append({"entity": ent, "remaining": cnt})
    # ترتيب اختياري: الأكثر بقاءً أولاً
    items.sort(key=lambda x: x["remaining"], reverse=True)
    return items

def load_assignments():
    data = load_json(ASSIGNMENTS_JSON, {})
    if not isinstance(data, dict):
        data = {}
    return data

def save_assignment(trainee_id: str, specialty: str, entity: str):
    assignments = load_assignments()
    trainee_id = str(trainee_id)

    assignments[trainee_id] = {
        "trainee_id": trainee_id,
        "specialty": specialty,
        "entity": entity,
        "timestamp": datetime.now().isoformat()
    }
    save_json(ASSIGNMENTS_JSON, assignments)

def render_docx_to_pdf(docx_path: str, out_dir: str):
    """
    تحويل DOCX -> PDF باستخدام LibreOffice (يعمل على Render)
    """
    # libreoffice command: soffice --headless --convert-to pdf --outdir OUTDIR DOCX
    subprocess.run(
        ["soffice", "--headless", "--nologo", "--nofirststartwizard", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
        check=True
    )
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf_path):
        raise RuntimeError("فشل إنشاء PDF")
    return pdf_path

def make_letter_docx(trainee: dict, entity: str, output_path: str):
    """
    تعبئة قالب Word placeholders باستخدام docxtpl
    """
    from docxtpl import DocxTemplate

    tpl = DocxTemplate(TEMPLATE_DOCX)

    # غيّر المفاتيح هنا إذا كانت حقول القالب عندك مختلفة
    context = {
        "TraineeName": trainee.get("اسم المتدرب", ""),
        "AcademicID": trainee.get("رقم المتدرب", ""),
        "Phone": trainee.get("رقم الجوال_norm", ""),
        "Specialty": trainee.get("التخصص", ""),
        "Program": trainee.get("البرنامج", ""),
        "Company": entity,
        "LetterNo": trainee.get("الرقم المرجعي", ""),  # إذا عندك عمود آخر عدّل
        "college_supervisor": trainee.get("المدرب", ""), # إذا عندك عمود آخر عدّل
        "course_ref": trainee.get("الرقم المرجعي", ""),  # إذا عندك عمود آخر عدّل
        "training_entity": entity,
    }

    tpl.render(context)
    tpl.save(output_path)

@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        try:
            trainee_id = (request.form.get("trainee_id") or "").strip()
            last4 = (request.form.get("last4") or "").strip()

            if not trainee_id or not last4:
                return render_template("index.html", error="فضلاً أدخل رقم المتدرب وآخر 4 أرقام من الجوال")

            df = read_students()

            matches = df[df["رقم المتدرب"] == str(trainee_id)]
            if matches.empty:
                return render_template("index.html", error="لم يتم العثور على رقم المتدرب")

            trainee = matches.iloc[0].to_dict()

            phone = trainee.get("رقم الجوال_norm", "")
            if len(phone) < 4 or phone[-4:] != last4:
                return render_template("index.html", error="آخر 4 أرقام من الجوال غير صحيحة")

            return redirect(url_for("dashboard", trainee_id=trainee_id))

        except Exception as e:
            error = str(e)

    return render_template("index.html", error=error)

@app.route("/dashboard/<trainee_id>", methods=["GET"])
def dashboard(trainee_id):
    error = None
    ok = None
    try:
        df = read_students()
        matches = df[df["رقم المتدرب"] == str(trainee_id)]
        if matches.empty:
            return redirect(url_for("index"))

        trainee = matches.iloc[0].to_dict()

        slots = ensure_slots(df)
        specialty = trainee.get("التخصص", "")
        total = sum(int(v) for v in slots.get(specialty, {}).values())
        remaining = sum(int(v) for v in slots.get(specialty, {}).values())
        entities = get_entities_for_specialty(slots, specialty)

        return render_template(
            "dashboard.html",
            trainee=trainee,
            total_slots=total,
            remaining_slots=remaining,
            entities=entities,
            error=error,
            ok=ok
        )
    except Exception as e:
        return render_template("dashboard.html", trainee={"رقم المتدرب": trainee_id, "اسم المتدرب": "", "التخصص": "", "البرنامج": ""}, total_slots=0, remaining_slots=0, entities=[], error=str(e), ok=None)

@app.route("/choose", methods=["POST"])
def choose():
    error = None
    ok = None
    try:
        action = (request.form.get("action") or "save").strip()
        trainee_id = (request.form.get("trainee_id") or "").strip()
        entity = (request.form.get("entity") or "").strip()

        if not trainee_id:
            return redirect(url_for("index"))

        df = read_students()
        matches = df[df["رقم المتدرب"] == str(trainee_id)]
        if matches.empty:
            return redirect(url_for("index"))
        trainee = matches.iloc[0].to_dict()
        specialty = trainee.get("التخصص", "")

        slots = ensure_slots(df)

        if not entity:
            error = "اختر جهة التدريب أولاً"
        else:
            # تحقق من وجود مقاعد
            spec_slots = slots.get(specialty, {})
            remaining = int(spec_slots.get(entity, 0))
            if remaining <= 0:
                error = "هذه الجهة لا يوجد لها مقاعد متبقية لتخصصك"
            else:
                # أنقص المقعد
                spec_slots[entity] = remaining - 1
                slots[specialty] = spec_slots
                save_json(SLOTS_JSON, slots)

                # احفظ اختيار المتدرب (بدون append)
                save_assignment(trainee_id=str(trainee_id), specialty=specialty, entity=entity)
                ok = "تم حفظ اختيارك بنجاح"

                if action == "pdf":
                    # توليد DOCX ثم PDF
                    with tempfile.TemporaryDirectory() as td:
                        docx_out = os.path.join(td, f"letter_{trainee_id}.docx")
                        make_letter_docx(trainee, entity, docx_out)
                        pdf_path = render_docx_to_pdf(docx_out, td)
                        return send_file(pdf_path, as_attachment=True, download_name=f"letter_{trainee_id}.pdf")

        # إعادة عرض اللوحة
        total = sum(int(v) for v in slots.get(specialty, {}).values())
        remaining_total = sum(int(v) for v in slots.get(specialty, {}).values())
        entities_list = get_entities_for_specialty(slots, specialty)

        return render_template(
            "dashboard.html",
            trainee=trainee,
            total_slots=total,
            remaining_slots=remaining_total,
            entities=entities_list,
            error=error,
            ok=ok
        )

    except Exception as e:
        return redirect(url_for("index") + f"?err={str(e)}")

if __name__ == "__main__":
    app.run(debug=True)

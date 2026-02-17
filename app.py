import os
import json
import uuid
import shutil
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for

# =========================
# إعدادات عامة
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

STUDENTS_XLSX = DATA_DIR / "students.xlsx"
ASSIGNMENTS_JSON = DATA_DIR / "assignments.json"
TEMPLATE_DOCX = DATA_DIR / "letter_template.docx"

OUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20MB


# =========================
# أدوات مساعدة
# =========================
def _load_json(path: Path, default):
    if not path.exists():
        _save_json(path, default)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        _save_json(path, default)
        return default


def _save_json(path: Path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _digits_only(x) -> str:
    return "".join([ch for ch in str(x) if ch.isdigit()])


def _read_students_df() -> pd.DataFrame:
    if not STUDENTS_XLSX.exists():
        raise FileNotFoundError(f"الملف غير موجود: {STUDENTS_XLSX}")

    df = pd.read_excel(STUDENTS_XLSX)

    required_cols = ["رقم المتدرب", "اسم المتدرب", "رقم الجوال", "جهة التدريب", "برنامج"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"أعمدة ناقصة في students.xlsx: {missing}")

    # تنظيف وتحويل
    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()
    df["برنامج"] = df["برنامج"].astype(str).str.strip()

    # رقم الجوال قد يأتي بصيغ علمية أو فيها رموز
    df["رقم الجوال"] = df["رقم الجوال"].apply(_digits_only)

    return df


def _get_student(training_id: str):
    df = _read_students_df()
    row = df[df["رقم المتدرب"] == str(training_id).strip()]
    if row.empty:
        return None
    return row.iloc[0].to_dict()


def _get_entities_capacity_from_excel() -> dict:
    """
    يحسب عدد الفرص لكل جهة من ملف Excel:
    عدد الفرص = عدد مرات تكرار الجهة في عمود "جهة التدريب"
    """
    df = _read_students_df()
    counts = df["جهة التدريب"].dropna().astype(str).str.strip().value_counts().to_dict()
    # تأكد أنها int
    return {k: int(v) for k, v in counts.items() if str(k).strip()}


def _count_used_slots_by_entity() -> dict:
    """
    يحسب عدد مرات الحجز/الطباعة لكل جهة من assignments.json
    """
    assignments = _load_json(ASSIGNMENTS_JSON, default=[])
    used = {}
    for a in assignments:
        ent = str(a.get("entity", "")).strip()
        if not ent:
            continue
        used[ent] = used.get(ent, 0) + 1
    return used


def _get_available_entities():
    """
    يرجع قائمة الجهات المتاحة فقط:
    remaining = capacity_from_excel - used_count
    يعرض فقط remaining > 0
    """
    capacity = _get_entities_capacity_from_excel()
    used = _count_used_slots_by_entity()

    available = []
    for ent, mx in capacity.items():
        remaining = int(mx) - int(used.get(ent, 0))
        if remaining > 0:
            available.append((ent, remaining, int(mx)))

    # ترتيب تنازلي حسب المتبقي (اختياري)
    available.sort(key=lambda x: x[1], reverse=True)
    return available


def _reserve_slot(entity_name: str) -> bool:
    """
    يحجز فرصة (يعني يسجل assignment) بشرط أن remaining > 0 الآن
    """
    entity_name = str(entity_name).strip()
    if not entity_name:
        return False

    capacity = _get_entities_capacity_from_excel()
    if entity_name not in capacity:
        return False

    used = _count_used_slots_by_entity()
    remaining = int(capacity[entity_name]) - int(used.get(entity_name, 0))
    if remaining <= 0:
        return False

    return True  # الحجز الفعلي يتم عند إضافة assignment


def _add_assignment(training_id: str, entity_name: str):
    assignments = _load_json(ASSIGNMENTS_JSON, default=[])
    assignments.append(
        {
            "training_id": str(training_id),
            "entity": str(entity_name).strip(),
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }
    )
    _save_json(ASSIGNMENTS_JSON, assignments)


def _render_docx_from_template(output_docx: Path, student: dict, entity_name: str):
    """
    قالب DOCX يحتوي placeholders مثل:
    {{اسم المتدرب}} {{رقم المتدرب}} {{رقم الجوال}} {{جهة التدريب}} {{برنامج}} {{التاريخ}}
    """
    try:
        from docx import Document
    except Exception as e:
        raise RuntimeError("مكتبة python-docx غير متوفرة. تأكد أنها داخل requirements.txt") from e

    if not TEMPLATE_DOCX.exists():
        raise FileNotFoundError(f"قالب الخطاب غير موجود: {TEMPLATE_DOCX}")

    doc = Document(str(TEMPLATE_DOCX))

    repl = {
        "{{اسم المتدرب}}": str(student.get("اسم المتدرب", "")),
        "{{رقم المتدرب}}": str(student.get("رقم المتدرب", "")),
        "{{رقم الجوال}}": str(student.get("رقم الجوال", "")),
        "{{جهة التدريب}}": str(entity_name),
        "{{برنامج}}": str(student.get("برنامج", "")),
        "{{التاريخ}}": datetime.now().strftime("%Y/%m/%d"),
    }

    def replace_in_paragraph(paragraph):
        for k, v in repl.items():
            if k in paragraph.text:
                for run in paragraph.runs:
                    if k in run.text:
                        run.text = run.text.replace(k, v)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)

    doc.save(str(output_docx))


def _convert_docx_to_pdf(input_docx: Path, output_pdf: Path):
    if output_pdf.exists():
        output_pdf.unlink()

    outdir = output_pdf.parent
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(outdir),
        str(input_docx),
    ]
    subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    produced = outdir / (input_docx.stem + ".pdf")
    if not produced.exists():
        raise RuntimeError("فشل تحويل DOCX إلى PDF. لم يتم العثور على ملف PDF الناتج.")

    if produced != output_pdf:
        shutil.move(str(produced), str(output_pdf))


# =========================
# صفحات HTML
# =========================
PAGE_LOGIN = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0;}
    .top-image{width:100%;height:25vh;overflow:hidden;background:#fff;}
    .top-image img{width:100%;height:100%;object-fit:cover;display:block;}
    .container{
      max-width:900px;margin:30px auto;background:#fff;padding:40px;border-radius:16px;
      box-shadow:0 10px 30px rgba(0,0,0,.08);text-align:center;
    }
    h1{margin:0 0 10px 0;font-size:44px;}
    p{color:#444;margin:10px 0 25px 0;}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:10px 0 18px 0;text-align:right;}
    label{display:block;margin-bottom:8px;font-weight:bold;}
    input{
      width:100%;padding:14px 16px;border:1px solid #ddd;border-radius:12px;font-size:18px;outline:none;
    }
    button{
      width:100%;padding:16px;background:#0f172a;color:#fff;border:none;border-radius:14px;font-size:20px;
      cursor:pointer;margin-top:10px;
    }
    .error{color:#b91c1c;margin-top:14px;font-weight:bold;}
    .note{color:#666;margin-top:12px;font-size:14px;}
  </style>
</head>
<body>

  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="container">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <p>يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

    <form method="POST">
      <div class="grid">
        <div>
          <label>الرقم التدريبي</label>
          <input name="training_id" required placeholder="مثال: 444229747" inputmode="numeric">
        </div>
        <div>
          <label>آخر 4 أرقام من الجوال</label>
          <input name="phone_last4" required maxlength="4" placeholder="مثال: 6101" inputmode="numeric">
        </div>
      </div>
      <button type="submit">دخول</button>
    </form>

    {% if error %}
      <div class="error">{{error}}</div>
    {% endif %}

    <div class="note">
      يتم احتساب الفرص لكل جهة من عدد تكرارها داخل ملف <b>students.xlsx</b>.
    </div>
  </div>
</body>
</html>
"""

PAGE_DASHBOARD = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0;}
    .top-image{width:100%;height:25vh;overflow:hidden;background:#fff;}
    .top-image img{width:100%;height:100%;object-fit:cover;display:block;}
    .container{
      max-width:980px;margin:30px auto;background:#fff;padding:30px;border-radius:16px;
      box-shadow:0 10px 30px rgba(0,0,0,.08);
    }
    h1{margin:0 0 10px 0;font-size:40px;text-align:center;}
    .card{background:#f1f5ff;padding:18px;border-radius:14px;margin:18px 0;line-height:1.9;}
    select{
      width:100%;padding:14px 16px;border:1px solid #ddd;border-radius:12px;font-size:18px;outline:none;background:#fff;
      margin-top:10px;
    }
    button{
      width:100%;padding:16px;background:#0f172a;color:#fff;border:none;border-radius:14px;font-size:20px;
      cursor:pointer;margin-top:14px;
    }
    .error{color:#b91c1c;margin-top:10px;font-weight:bold;text-align:center;}
    .slots{color:#444;font-size:14px;margin-top:8px;}
    .back{display:block;text-align:center;margin-top:14px;color:#0f172a;text-decoration:none;}
  </style>
</head>
<body>

  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="container">
    <h1>{{title}}</h1>

    <div class="card">
      <div><b>الاسم:</b> {{student["اسم المتدرب"]}}</div>
      <div><b>الرقم التدريبي:</b> {{student["رقم المتدرب"]}}</div>
      <div><b>البرنامج:</b> {{student["برنامج"]}}</div>
      <div><b>الجوال:</b> {{student["رقم الجوال"]}}</div>
    </div>

    <form method="POST" action="{{url_for('generate_pdf', training_id=student['رقم المتدرب'])}}">
      <b>اختر جهة التدريب المتاحة:</b>
      <select name="entity" required>
        <option value="" disabled selected>اختر جهة...</option>
        {% for ent, remaining, mx in entities %}
          <option value="{{ent}}">{{ent}} (متبقي {{remaining}} من {{mx}})</option>
        {% endfor %}
      </select>
      <div class="slots">
        الفرص لكل جهة = عدد تكرار الجهة داخل ملف Excel. الجهات التي انتهت (0) تختفي تلقائياً.
      </div>

      <button type="submit">طباعة خطاب التوجيه PDF</button>
    </form>

    {% if error %}
      <div class="error">{{error}}</div>
    {% endif %}

    <a class="back" href="{{url_for('logout')}}">تسجيل خروج</a>
  </div>

</body>
</html>
"""


# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
def login():
    error = None

    if request.method == "POST":
        training_id = (request.form.get("training_id") or "").strip()
        phone_last4 = (request.form.get("phone_last4") or "").strip()

        if not training_id or not phone_last4 or len(phone_last4) != 4:
            error = "الرجاء إدخال الرقم التدريبي وآخر 4 أرقام من الجوال بشكل صحيح."
            return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error)

        try:
            student = _get_student(training_id)
            if not student:
                error = "بيانات الدخول غير صحيحة."
                return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error)

            full_phone = _digits_only(student.get("رقم الجوال", ""))
            if len(full_phone) < 4 or full_phone[-4:] != phone_last4:
                error = "بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال."
                return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error)

            return redirect(url_for("dashboard", training_id=training_id))

        except Exception as e:
            error = f"حدث خطأ أثناء التحقق: {str(e)}"

    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=error)


@app.route("/dashboard/<training_id>", methods=["GET"])
def dashboard(training_id):
    try:
        student = _get_student(training_id)
        if not student:
            return redirect(url_for("login"))

        entities = _get_available_entities()
        error = None
        if not entities:
            error = "لا توجد جهات متاحة حالياً (انتهت جميع الفرص)."

        return render_template_string(
            PAGE_DASHBOARD,
            title=APP_TITLE,
            student=student,
            entities=entities,
            error=error,
        )
    except Exception as e:
        return f"Internal Server Error: {str(e)}", 500


@app.route("/generate/<training_id>", methods=["POST"])
def generate_pdf(training_id):
    entity = (request.form.get("entity") or "").strip()
    if not entity:
        return redirect(url_for("dashboard", training_id=training_id))

    try:
        student = _get_student(training_id)
        if not student:
            return redirect(url_for("login"))

        # تحقق أن الفرصة ما زالت متاحة قبل الطباعة
        if not _reserve_slot(entity):
            entities = _get_available_entities()
            return render_template_string(
                PAGE_DASHBOARD,
                title=APP_TITLE,
                student=student,
                entities=entities,
                error="هذه الجهة لم تعد متاحة (انتهت الفرص). اختر جهة أخرى.",
            )

        # سجل الحجز (يعتبر خصم فرصة)
        _add_assignment(training_id, entity)

        token = uuid.uuid4().hex
        out_docx = OUT_DIR / f"letter_{training_id}_{token}.docx"
        out_pdf = OUT_DIR / f"letter_{training_id}_{token}.pdf"

        _render_docx_from_template(out_docx, student, entity)
        _convert_docx_to_pdf(out_docx, out_pdf)

        # تنظيف docx
        try:
            out_docx.unlink(missing_ok=True)
        except Exception:
            pass

        return send_file(
            out_pdf,
            as_attachment=True,
            download_name=f"خطاب_التوجيه_{training_id}.pdf",
            mimetype="application/pdf",
        )

    except subprocess.CalledProcessError:
        return "خطأ أثناء تحويل PDF. تأكد أن LibreOffice (soffice) مثبت داخل Dockerfile.", 500
    except Exception as e:
        return f"Internal Server Error: {str(e)}", 500


@app.route("/logout")
def logout():
    return redirect(url_for("login"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port, debug=False)

import os
import re
import json
import subprocess
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, render_template_string, abort

from docxtpl import DocxTemplate

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"

DATA_DIR.mkdir(exist_ok=True)
STATIC_DIR.mkdir(exist_ok=True)

STUDENTS_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"
TEMPLATE_DOCX = DATA_DIR / "letter_template.docx"
OUTPUT_DIR = DATA_DIR / "out"
OUTPUT_DIR.mkdir(exist_ok=True)


# -----------------------------
# Helpers
# -----------------------------
def _norm_col(s: str) -> str:
    """Normalize Arabic/spacing variations in column names."""
    s = str(s).strip()
    s = s.replace("إسم", "اسم")
    s = re.sub(r"\s+", " ", s)
    return s


def load_students_df() -> pd.DataFrame:
    if not STUDENTS_FILE.exists():
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {STUDENTS_FILE}")

    df = pd.read_excel(STUDENTS_FILE)
    df.columns = [_norm_col(c) for c in df.columns]

    # Normalize required columns for your uploaded file
    required = [
        "رقم المتدرب",
        "اسم المتدرب",
        "رقم الجوال",
        "التخصص",
        "برنامج",
        "جهة التدريب",
        "المدرب",
        "الرقم المرجعي",
        "اسم المقرر",
        "القسم",
    ]
    # Handle: your file has "جهة التدريب " with trailing space -> normalized to "جهة التدريب"
    # Handle: your file has "إسم المتدرب" -> normalized to "اسم المتدرب"

    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"لم أجد الأعمدة المطلوبة: {missing}. الأعمدة الموجودة: {list(df.columns)}")

    # Clean values
    df["رقم المتدرب"] = df["رقم المتدرب"].astype(str).str.strip()
    df["رقم الجوال"] = df["رقم الجوال"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["التخصص"] = df["التخصص"].astype(str).str.strip()
    df["برنامج"] = df["برنامج"].astype(str).str.strip()
    df["جهة التدريب"] = df["جهة التدريب"].astype(str).str.strip()
    df["اسم المتدرب"] = df["اسم المتدرب"].astype(str).str.strip()
    df["المدرب"] = df["المدرب"].astype(str).str.strip()
    df["الرقم المرجعي"] = df["الرقم المرجعي"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["اسم المقرر"] = df["اسم المقرر"].astype(str).str.strip()
    df["القسم"] = df["القسم"].astype(str).str.strip()

    return df


def load_assignments() -> dict:
    if ASSIGNMENTS_FILE.exists():
        try:
            return json.loads(ASSIGNMENTS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_assignments(data: dict) -> None:
    ASSIGNMENTS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def last4(phone: str) -> str:
    digits = re.sub(r"\D", "", str(phone))
    return digits[-4:] if len(digits) >= 4 else digits


def compute_total_slots(df: pd.DataFrame) -> dict:
    """
    الفرص = عدد تكرار (جهة التدريب) داخل نفس (التخصص) في ملف students.xlsx
    """
    g = df.groupby(["التخصص", "جهة التدريب"]).size().reset_index(name="total")
    totals = {}
    for _, row in g.iterrows():
        spec = row["التخصص"]
        ent = row["جهة التدريب"]
        totals.setdefault(spec, {})[ent] = int(row["total"])
    return totals


def compute_remaining_slots(df: pd.DataFrame, assignments: dict) -> dict:
    totals = compute_total_slots(df)

    used = {}
    for tid, info in assignments.items():
        spec = info.get("التخصص")
        ent = info.get("جهة التدريب")
        if spec and ent:
            used.setdefault(spec, {})
            used[spec][ent] = used[spec].get(ent, 0) + 1

    remaining = {}
    for spec, ents in totals.items():
        remaining[spec] = {}
        for ent, total in ents.items():
            u = used.get(spec, {}).get(ent, 0)
            remaining[spec][ent] = max(total - u, 0)
    return remaining


def find_student(df: pd.DataFrame, trainee_id: str, phone_last4: str):
    trainee_id = str(trainee_id).strip()
    phone_last4 = str(phone_last4).strip()

    sub = df[df["رقم المتدرب"] == trainee_id]
    if sub.empty:
        return None

    # match last4
    sub = sub[sub["رقم الجوال"].apply(lambda x: last4(x) == phone_last4)]
    if sub.empty:
        return None

    # student may appear multiple rows, take first row
    row = sub.iloc[0].to_dict()
    return row


def render_letter_docx_and_pdf(context: dict, out_basename: str):
    """
    Fill data/letter_template.docx using docxtpl placeholders, then convert to PDF using LibreOffice (soffice).
    """
    if not TEMPLATE_DOCX.exists():
        raise FileNotFoundError(f"قالب الوورد غير موجود: {TEMPLATE_DOCX}")

    out_docx = OUTPUT_DIR / f"{out_basename}.docx"
    out_pdf = OUTPUT_DIR / f"{out_basename}.pdf"

    doc = DocxTemplate(str(TEMPLATE_DOCX))
    doc.render(context)
    doc.save(str(out_docx))

    # Convert to PDF with LibreOffice (installed in Dockerfile)
    # soffice path usually exists after install: /usr/bin/soffice
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(OUTPUT_DIR),
        str(out_docx),
    ]
    p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    if p.returncode != 0 or not out_pdf.exists():
        raise RuntimeError(f"فشل تحويل PDF. stderr: {p.stderr[:800]}")

    return out_docx, out_pdf


# -----------------------------
# HTML (inline templates)
# -----------------------------
BASE_CSS = """
<style>
  body{margin:0;background:#f4f6fb;font-family:Tahoma,Arial;direction:rtl;}
  .top-image img{width:100%;height:auto;display:block;max-height:220px;object-fit:contain;background:#fff;}
  .wrap{max-width:1100px;margin:20px auto;padding:0 14px;}
  .card{background:#fff;border-radius:22px;box-shadow:0 10px 30px rgba(0,0,0,.06);padding:24px;}
  h1{margin:0 0 6px 0;font-size:44px;text-align:center;}
  .sub{margin:0 0 18px 0;text-align:center;color:#444;font-size:18px;}
  .row{display:flex;gap:16px;flex-wrap:wrap;justify-content:space-between;margin-top:10px;}
  .field{flex:1;min-width:260px}
  label{display:block;margin:10px 0 8px 0;font-weight:700;font-size:18px}
  input,select{width:100%;padding:14px 16px;border-radius:16px;border:1px solid #ddd;font-size:18px}
  .btn{width:100%;border:none;border-radius:20px;padding:16px 18px;background:#0b1730;color:#fff;
       font-size:22px;font-weight:700;cursor:pointer;margin-top:14px}
  .error{color:#c60000;font-weight:700;margin-top:12px;white-space:pre-wrap}
  .hint{color:#666;margin-top:8px}
  .pillrow{display:flex;gap:10px;flex-wrap:wrap;justify-content:center;margin:14px 0 6px 0}
  .pill{background:#eef4ff;border-radius:999px;padding:10px 16px;font-weight:700}
  ul{margin:10px 0 0 0}
</style>
"""

HOME_HTML = """
<!doctype html><html lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>بوابة خطاب التوجيه - التدريب التعاوني</title>
""" + BASE_CSS + """
</head><body>
<div class="top-image"><img src="/static/header.jpg" alt="Header"></div>
<div class="wrap">
  <div class="card">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <p class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

    <form method="post" action="/">
      <div class="row">
        <div class="field">
          <label>الرقم التدريبي</label>
          <input name="trainee_id" placeholder="مثال: 444242291" required>
        </div>
        <div class="field">
          <label>آخر 4 أرقام من الجوال</label>
          <input name="phone_last4" placeholder="مثال: 5513" required>
        </div>
      </div>
      <button class="btn" type="submit">دخول</button>
    </form>

    {% if error %}
      <div class="error">{{error}}</div>
    {% endif %}

    <div class="hint">ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b></div>
  </div>
</div>
</body></html>
"""

CHOOSE_HTML = """
<!doctype html><html lang="ar"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>اختيار جهة التدريب</title>
""" + BASE_CSS + """
</head><body>
<div class="top-image"><img src="/static/header.jpg" alt="Header"></div>
<div class="wrap">
  <div class="card">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <p class="sub">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

    <div class="pillrow">
      <div class="pill">المتدرب: {{name}}</div>
      <div class="pill">رقم المتدرب: {{tid}}</div>
      <div class="pill">التخصص/البرنامج: {{spec}} — {{program}}</div>
    </div>

    {% if already %}
      <div style="margin-top:12px;font-weight:700">
        تم تسجيل اختيارك مسبقاً: <span style="color:#0b1730">{{already}}</span>
      </div>
      <div style="margin-top:8px">
        <a href="/letter?tid={{tid}}" target="_blank">تحميل/طباعة خطاب التوجيه PDF</a>
      </div>
    {% else %}
      <form method="post" action="/choose?tid={{tid}}">
        <label>جهة التدريب المتاحة</label>
        <select name="entity" required>
          <option value="">اختر الجهة...</option>
          {% for ent, rem in options %}
            <option value="{{ent}}">{{ent}} — (متبقي: {{rem}})</option>
          {% endfor %}
        </select>
        <button class="btn" type="submit">حفظ الاختيار</button>
      </form>
    {% endif %}

    {% if msg %}
      <div class="error" style="color:#0a6a00">{{msg}}</div>
    {% endif %}
    {% if error %}
      <div class="error">{{error}}</div>
    {% endif %}

    <hr style="margin:18px 0;border:none;border-top:1px solid #eee">

    <div style="font-weight:800;font-size:20px">ملخص الفرص المتبقية حسب الجهات داخل تخصصك:</div>
    <ul>
      {% for ent, rem in summary %}
        <li>{{ent}} : <b>{{rem}}</b> فرصة</li>
      {% endfor %}
    </ul>
  </div>
</div>
</body></html>
"""


# -----------------------------
# Routes
# -----------------------------
@app.get("/")
def home_get():
    return render_template_string(HOME_HTML, error=None)


@app.post("/")
def home_post():
    try:
        df = load_students_df()
        trainee_id = request.form.get("trainee_id", "").strip()
        phone_last4 = request.form.get("phone_last4", "").strip()

        stu = find_student(df, trainee_id, phone_last4)
        if not stu:
            return render_template_string(HOME_HTML, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

        return redirect(url_for("choose_get", tid=trainee_id))
    except Exception as e:
        return render_template_string(HOME_HTML, error=f"حدث خطأ أثناء التحميل/التحقق: {e}")


@app.get("/choose")
def choose_get():
    tid = request.args.get("tid", "").strip()
    if not tid:
        return redirect(url_for("home_get"))

    try:
        df = load_students_df()
        # show student basic info (first row by trainee id)
        row = df[df["رقم المتدرب"] == tid]
        if row.empty:
            return redirect(url_for("home_get"))

        r0 = row.iloc[0].to_dict()
        name = r0["اسم المتدرب"]
        spec = r0["التخصص"]
        program = r0["برنامج"]

        assignments = load_assignments()
        already = None
        if tid in assignments:
            already = assignments[tid].get("جهة التدريب")

        remaining = compute_remaining_slots(df, assignments)
        rem_for_spec = remaining.get(spec, {})

        # options only with remaining > 0
        options = sorted([(ent, rem) for ent, rem in rem_for_spec.items() if rem > 0], key=lambda x: (-x[1], x[0]))
        summary = sorted([(ent, rem) for ent, rem in rem_for_spec.items()], key=lambda x: (-x[1], x[0]))

        if not options and not already:
            error = "لا توجد جهات متاحة حالياً لهذا التخصص."
        else:
            error = None

        return render_template_string(
            CHOOSE_HTML,
            tid=tid,
            name=name,
            spec=spec,
            program=program,
            options=options,
            summary=summary,
            already=already,
            error=error,
            msg=None,
        )
    except Exception as e:
        return render_template_string(CHOOSE_HTML, tid=tid, name="", spec="", program="", options=[], summary=[], already=None, error=str(e), msg=None)


@app.post("/choose")
def choose_post():
    tid = request.args.get("tid", "").strip()
    if not tid:
        return redirect(url_for("home_get"))

    try:
        df = load_students_df()
        row = df[df["رقم المتدرب"] == tid]
        if row.empty:
            return redirect(url_for("home_get"))
        r0 = row.iloc[0].to_dict()
        spec = r0["التخصص"]

        entity = request.form.get("entity", "").strip()
        if not entity:
            return redirect(url_for("choose_get", tid=tid))

        assignments = load_assignments()
        # prevent re-choose (you can remove this if you want allow change)
        if tid in assignments:
            return redirect(url_for("choose_get", tid=tid))

        remaining = compute_remaining_slots(df, assignments)
        rem = remaining.get(spec, {}).get(entity, 0)
        if rem <= 0:
            return render_template_string(CHOOSE_HTML, tid=tid, name=r0["اسم المتدرب"], spec=spec, program=r0["برنامج"],
                                          options=[], summary=sorted(list(remaining.get(spec, {}).items()), key=lambda x: (-x[1], x[0])),
                                          already=None, error="هذه الجهة لم تعد متاحة (انتهت الفرص). اختر جهة أخرى.", msg=None)

        assignments[tid] = {
            "رقم المتدرب": tid,
            "اسم المتدرب": r0["اسم المتدرب"],
            "رقم الجوال": r0["رقم الجوال"],
            "القسم": r0["القسم"],
            "التخصص": spec,
            "برنامج": r0["برنامج"],
            "المدرب": r0["المدرب"],
            "الرقم المرجعي": r0["الرقم المرجعي"],
            "اسم المقرر": r0["اسم المقرر"],
            "جهة التدريب": entity,
            "timestamp": datetime.utcnow().isoformat(),
        }
        save_assignments(assignments)
        return redirect(url_for("choose_get", tid=tid))

    except Exception as e:
        return render_template_string(CHOOSE_HTML, tid=tid, name="", spec="", program="", options=[], summary=[], already=None, error=str(e), msg=None)


@app.get("/letter")
def letter_pdf():
    tid = request.args.get("tid", "").strip()
    if not tid:
        return redirect(url_for("home_get"))

    assignments = load_assignments()
    if tid not in assignments:
        return "لم يتم العثور على اختيار محفوظ لهذا المتدرب.", 404

    info = assignments[tid]

    # تجهيز بيانات القالب (Placeholders في الوورد)
    context = {
        "trainee_id": info.get("رقم المتدرب", ""),
        "trainee_name": info.get("اسم المتدرب", ""),
        "phone": info.get("رقم الجوال", ""),
        "department": info.get("القسم", ""),
        "specialization": info.get("التخصص", ""),
        "program": info.get("برنامج", ""),
        "trainer": info.get("المدرب", ""),
        "course_ref": info.get("الرقم المرجعي", ""),
        "course_name": info.get("اسم المقرر", ""),
        "training_entity": info.get("جهة التدريب", ""),
        # عدّلها إذا عندك اسم مشرف ثابت
        "college_supervisor": "اسم المشرف من الكلية",
        "today_hijri": "",  # إذا تبي هجري لاحقاً نضيفها
        "today_greg": datetime.now().strftime("%Y-%m-%d"),
    }

    try:
        out_base = f"letter_{tid}"
        _, out_pdf = render_letter_docx_and_pdf(context, out_base)
        return send_file(out_pdf, as_attachment=True, download_name=f"خطاب_توجيه_{tid}.pdf")
    except Exception as e:
        return f"حدث خطأ أثناء إنشاء PDF: {e}", 500


# Local run (optional)
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

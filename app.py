قام بالتفكير لمدة 1m 22s

تمام — مشكلتك الآن لها 3 أسباب رئيسية:

ملف الإكسل عندك أعمدته عربية وبأسماء مختلفة (مثل إسم المتدرب وليس اسم المتدرب، و جهة التدريب عندك فيها مسافة آخر الاسم: جهة التدريب ). لذلك كودك كان يبحث عن أعمدة غير موجودة فيعطي خطأ.

طباعة الخطاب كانت تعطي FileNotFoundError: soffice لأن Render غالبًا يشغل مشروعك كـ Python وليس Docker، وبالتالي LibreOffice غير موجود. الحل الأسلم: نُخرج PDF مباشرة من بايثون بدون soffice.

“الفرص” لازم تُحسب من نفس ملف الإكسل: كل جهة تدريبية مكررة داخل نفس التخصص = عدد فرصها، وتتناقص مع كل اختيار، وتختفي إذا انتهت.

أعطيك الآن كود كامل من البداية للنهاية لملف app.py فقط (انسخه والصقه كامل)، وهو متوافق مع ملفك data/students.xlsx الذي أعمدته مثل:
رقم المتدرب ، رقم الجوال ، إسم المتدرب ، التخصص ، برنامج ، جهة التدريب … إلخ.

1) تأكد من مكان الملف والصورة

لازم يكون عندك في GitHub بهذا الشكل:

data/students.xlsx

static/header.jpg

عندك الآن مجلد static ممتاز، وارفع داخله header.jpg
وعندك مجلد data ممتاز، وارفع داخله students.xlsx

2) استبدل ملف app.py بالكامل بهذا الكود
import os
import json
import re
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, render_template_string, abort

# PDF (بدون soffice)
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# لتشكيل العربية في PDF (مهم)
import arabic_reshaper
from bidi.algorithm import get_display


APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

DATA_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"
OUT_DIR.mkdir(exist_ok=True)


def _norm_col(s: str) -> str:
    """Normalize column names: remove spaces, unify Alef forms, etc."""
    s = str(s).strip()
    s = s.replace("\u200f", "").replace("\u200e", "")
    s = re.sub(r"\s+", " ", s)
    return s


def _norm_text(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u200f", "").replace("\u200e", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _digits_only(s) -> str:
    """Keep digits only, handle floats from Excel."""
    if s is None:
        return ""
    s = str(s).strip()
    # لو جاء مثل 535785513.0
    s = s.replace(".0", "")
    s = re.sub(r"\D+", "", s)
    return s


def load_students() -> pd.DataFrame:
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"ملف الطلاب غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE)

    # Normalize column names (خصوصاً وجود مسافة في "جهة التدريب ")
    df.columns = [_norm_col(c) for c in df.columns]

    # أعمدة لازمة في ملفك حسب الإكسل المرفوع
    required_any = {
        "رقم المتدرب": ["رقم المتدرب", "رقم_المتدرب", "Trainee ID", "trainee_id"],
        "رقم الجوال": ["رقم الجوال", "رقم_الجوال", "الجوال", "Mobile", "mobile"],
        "إسم المتدرب": ["إسم المتدرب", "اسم المتدرب", "اسم_المتدرب", "Name", "name"],
        "التخصص": ["التخصص", "تخصص", "Specialization", "specialization"],
        "برنامج": ["برنامج", "Program", "program"],
        "جهة التدريب": ["جهة التدريب", "جهة التدريب ", "جهة_التدريب", "Training Entity", "entity"],
    }

    # خريطة: المفتاح القياسي -> الاسم الحقيقي الموجود
    colmap = {}
    cols = set(df.columns)

    for std_name, options in required_any.items():
        found = None
        for opt in options:
            optn = _norm_col(opt)
            for real in cols:
                if _norm_col(real) == optn:
                    found = real
                    break
            if found:
                break
        if not found:
            raise KeyError(
                f"لم أجد عمود '{std_name}' داخل ملف students.xlsx.\n"
                f"الأعمدة الموجودة: {list(df.columns)}"
            )
        colmap[std_name] = found

    # Rename to standard names
    df = df.rename(columns={
        colmap["رقم المتدرب"]: "trainee_id",
        colmap["رقم الجوال"]: "mobile",
        colmap["إسم المتدرب"]: "trainee_name",
        colmap["التخصص"]: "specialization",
        colmap["برنامج"]: "program",
        colmap["جهة التدريب"]: "entity",
    })

    # تنظيف النصوص
    df["trainee_id"] = df["trainee_id"].apply(_digits_only)
    df["mobile"] = df["mobile"].apply(_digits_only)
    df["trainee_name"] = df["trainee_name"].apply(_norm_text)
    df["specialization"] = df["specialization"].apply(_norm_text)
    df["program"] = df["program"].apply(_norm_text)
    df["entity"] = df["entity"].apply(_norm_text)

    # حذف صفوف ناقصة
    df = df[(df["trainee_id"] != "") & (df["mobile"] != "") & (df["specialization"] != "")]
    return df


def load_assignments() -> dict:
    if ASSIGNMENTS_FILE.exists():
        try:
            return json.loads(ASSIGNMENTS_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_assignments(d: dict) -> None:
    ASSIGNMENTS_FILE.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")


def build_slots(df: pd.DataFrame) -> dict:
    """
    كل (تخصص + جهة) مكررة في ملف الإكسل = عدد الفرص المتاحة.
    """
    g = df.groupby(["specialization", "entity"]).size().reset_index(name="slots")
    slots = {}
    for _, r in g.iterrows():
        spec = r["specialization"]
        ent = r["entity"]
        n = int(r["slots"])
        if spec not in slots:
            slots[spec] = {}
        if ent:  # تجاهل الفارغ
            slots[spec][ent] = n
    return slots


def remaining_slots_for_spec(spec: str, slots: dict, assignments: dict) -> dict:
    """
    يحسب المتبقي لكل جهة داخل تخصص معيّن بناء على اختيارات الطلاب.
    """
    base = dict(slots.get(spec, {}))
    # احسب كم واحد اختار كل جهة داخل هذا التخصص
    used = {}
    for tid, a in assignments.items():
        if a.get("specialization") == spec:
            ent = a.get("entity")
            if ent:
                used[ent] = used.get(ent, 0) + 1

    remaining = {}
    for ent, total in base.items():
        rem = int(total) - int(used.get(ent, 0))
        if rem > 0:
            remaining[ent] = rem
    return remaining


def last4(mobile_digits: str) -> str:
    mobile_digits = _digits_only(mobile_digits)
    return mobile_digits[-4:] if len(mobile_digits) >= 4 else mobile_digits


# ---------- PDF helpers ----------
def register_ar_font() -> str:
    """
    يحاول تسجيل خط عربي موجود غالباً في لينكس.
    """
    candidates = [
        "/usr/share/fonts/truetype/noto/NotoNaskhArabic-Regular.ttf",
        "/usr/share/fonts/truetype/noto/NotoSansArabic-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            pdfmetrics.registerFont(TTFont("AR", p))
            return "AR"
    # fallback
    return "Helvetica"


def ar(text: str) -> str:
    """
    تشكيل العربية + اتجاه RTL ليظهر صحيح في PDF.
    """
    text = _norm_text(text)
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


FONT_NAME = register_ar_font()

# ---------- Flask ----------
app = Flask(__name__)


PAGE_LOGIN = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8" />
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{
      width:100%;
      height:25vh;
      overflow:hidden;
      background:#fff;
      display:flex;
      align-items:center;
      justify-content:center;
    }
    .top-image img{
      width:100%;
      height:100%;
      object-fit:contain; /* عشان ما تنقص من فوق وتحت */
      display:block;
    }

    .wrap{max-width:1100px;margin:18px auto;padding:0 16px;}
    .card{
      background:#fff;border-radius:22px;box-shadow:0 10px 30px rgba(0,0,0,.08);
      padding:28px;margin-top:18px;
    }
    h1{margin:0 0 12px 0;font-size:44px;text-align:center}
    .sub{text-align:center;color:#444;margin-bottom:22px}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:end}
    label{display:block;font-weight:700;margin-bottom:8px}
    input{
      width:100%;padding:16px 18px;border-radius:16px;border:1px solid #e2e2e2;font-size:18px;
    }
    .btn{
      width:100%;margin-top:18px;padding:18px;border-radius:18px;border:none;
      background:#0b1833;color:#fff;font-size:22px;cursor:pointer;
    }
    .err{color:#c00;text-align:center;margin-top:12px;font-weight:700}
    .hint{color:#666;text-align:center;margin-top:10px}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <div class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</div>

      <form method="POST" action="{{ url_for('login') }}">
        <div class="grid">
          <div>
            <label>الرقم التدريبي</label>
            <input name="trainee_id" placeholder="مثال: 444229747" required>
          </div>
          <div>
            <label>آخر 4 أرقام من الجوال</label>
            <input name="last4" placeholder="مثال: 6101" required>
          </div>
        </div>
        <button class="btn" type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}

      <div class="hint">ملاحظة: يتم قراءة ملف الطلاب من data/students.xlsx</div>
    </div>
  </div>
</body>
</html>
"""


PAGE_DASH = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8" />
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%;height:25vh;overflow:hidden;background:#fff;display:flex;align-items:center;justify-content:center;}
    .top-image img{width:100%;height:100%;object-fit:contain;display:block;}
    .wrap{max-width:1100px;margin:18px auto;padding:0 16px;}
    .card{background:#fff;border-radius:22px;box-shadow:0 10px 30px rgba(0,0,0,.08);padding:28px;margin-top:18px;}
    h1{margin:0 0 12px 0;font-size:40px;text-align:center}
    .chips{display:flex;gap:10px;justify-content:space-between;flex-wrap:wrap;margin:18px 0}
    .chip{background:#eef3ff;border-radius:999px;padding:10px 14px;font-weight:700}
    .row{display:grid;grid-template-columns:1fr;gap:12px;margin-top:10px}
    select{width:100%;padding:14px;border-radius:14px;border:1px solid #e2e2e2;font-size:18px;}
    .btn{width:100%;margin-top:14px;padding:16px;border-radius:18px;border:none;background:#0b1833;color:#fff;font-size:20px;cursor:pointer;}
    .err{color:#c00;text-align:center;margin-top:12px;font-weight:700}
    .ok{color:#0a7a2e;text-align:center;margin-top:12px;font-weight:700}
    ul{margin:10px 0 0 0;line-height:2}
    a{color:#0b1833}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>{{title}}</h1>
      <div style="text-align:center;color:#444;margin-bottom:12px;">
        اختر جهة التدريب المتاحة لتخصصك ثم اطبع خطاب التوجيه.
      </div>

      <div class="chips">
        <div class="chip">المتدرب: {{student.trainee_name}}</div>
        <div class="chip">رقم المتدرب: {{student.trainee_id}}</div>
        <div class="chip">التخصص: {{student.specialization}}</div>
        <div class="chip">البرنامج: {{student.program}}</div>
      </div>

      {% if chosen %}
        <div class="card" style="background:#f4f6f8">
          <div style="font-weight:800;margin-bottom:8px;">تم تسجيل اختيارك مسبقاً:</div>
          <div>الجهة المختارة: <b>{{chosen.entity}}</b></div>
          <div style="margin-top:10px;">
            <a href="{{ url_for('letter', tid=student.trainee_id) }}">تحميل/طباعة خطاب التوجيه PDF</a>
          </div>
        </div>
      {% endif %}

      <form method="POST" action="{{ url_for('choose') }}">
        <input type="hidden" name="tid" value="{{student.trainee_id}}">
        <div class="row">
          <label style="font-weight:800">جهة التدريب المتاحة (حسب تخصصك)</label>
          <select name="entity" required>
            <option value="">اختر الجهة...</option>
            {% for ent, rem in available.items() %}
              <option value="{{ent}}">{{ent}} (متبقي: {{rem}})</option>
            {% endfor %}
          </select>
        </div>
        <button class="btn" type="submit">حفظ الاختيار</button>
      </form>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}
      {% if message %}
        <div class="ok">{{message}}</div>
      {% endif %}

      <div class="card" style="background:#fff;margin-top:18px">
        <div style="font-weight:800;margin-bottom:8px;">ملخص الفرص المتبقية حسب الجهات داخل تخصصك:</div>
        {% if available %}
          <ul>
            {% for ent, rem in available.items() %}
              <li>{{ent}} : <b>{{rem}}</b> فرصة</li>
            {% endfor %}
          </ul>
        {% else %}
          <div class="err">لا توجد جهات متاحة حالياً لهذا التخصص.</div>
        {% endif %}
      </div>

    </div>
  </div>
</body>
</html>
"""


@app.get("/")
def index():
    return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=None)


@app.post("/")
def login():
    trainee_id = _digits_only(request.form.get("trainee_id"))
    l4 = _digits_only(request.form.get("last4"))

    try:
        df = load_students()
    except Exception as e:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=f"خطأ أثناء التحميل: {e}")

    row = df[df["trainee_id"] == trainee_id]
    if row.empty:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="الرقم التدريبي غير موجود.")

    mobile = str(row.iloc[0]["mobile"])
    if last4(mobile) != l4:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="آخر 4 أرقام من الجوال غير صحيحة.")

    return redirect(url_for("dashboard", tid=trainee_id))


@app.get("/dashboard")
def dashboard():
    tid = _digits_only(request.args.get("tid"))
    if not tid:
        return redirect(url_for("index"))

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("index"))
    student = row.iloc[0].to_dict()

    slots = build_slots(df)
    assignments = load_assignments()

    chosen = assignments.get(tid)
    available = remaining_slots_for_spec(student["specialization"], slots, assignments)

    return render_template_string(
        PAGE_DASH,
        title=APP_TITLE,
        student=student,
        chosen=chosen,
        available=available,
        error=None,
        message=None,
    )


@app.post("/choose")
def choose():
    tid = _digits_only(request.form.get("tid"))
    entity = _norm_text(request.form.get("entity"))

    if not tid or not entity:
        return redirect(url_for("dashboard", tid=tid))

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("index"))
    student = row.iloc[0].to_dict()

    slots = build_slots(df)
    assignments = load_assignments()

    # منع تغيير الاختيار إذا كان مسجل مسبقاً (اختياري: تقدر تسمح بالتغيير)
    if tid in assignments:
        return render_template_string(
            PAGE_DASH,
            title=APP_TITLE,
            student=student,
            chosen=assignments.get(tid),
            available=remaining_slots_for_spec(student["specialization"], slots, assignments),
            error="تم تسجيل اختيارك مسبقاً، لا يمكن التعديل حالياً.",
            message=None,
        )

    available = remaining_slots_for_spec(student["specialization"], slots, assignments)

    if entity not in available:
        return render_template_string(
            PAGE_DASH,
            title=APP_TITLE,
            student=student,
            chosen=None,
            available=available,
            error="هذه الجهة غير متاحة الآن (قد تكون انتهت الفرص أو ليست ضمن تخصصك).",
            message=None,
        )

    assignments[tid] = {
        "trainee_id": tid,
        "trainee_name": student["trainee_name"],
        "specialization": student["specialization"],
        "program": student["program"],
        "entity": entity,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
    }
    save_assignments(assignments)

    return redirect(url_for("dashboard", tid=tid))


@app.get("/letter")
def letter():
    tid = _digits_only(request.args.get("tid"))
    if not tid:
        abort(404)

    df = load_students()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        abort(404)
    student = row.iloc[0].to_dict()

    assignments = load_assignments()
    chosen = assignments.get(tid)
    if not chosen:
        return "لم يتم اختيار جهة تدريب بعد.", 400

    # توليد PDF مباشرة (بدون LibreOffice)
    pdf_path = OUT_DIR / f"letter_{tid}.pdf"

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    w, h = A4

    c.setTitle("خطاب التوجيه")

    # Header (نص فقط — والصورة موجودة في الصفحة الرئيسية)
    c.setFont(FONT_NAME, 16)
    c.drawRightString(w - 50, h - 60, ar("المؤسسة العامة للتدريب التقني والمهني"))
    c.setFont(FONT_NAME, 14)
    c.drawRightString(w - 50, h - 85, ar("الكلية التقنية بأبها - وحدة التدريب التعاوني"))

    # Body
    y = h - 140
    c.setFont(FONT_NAME, 18)
    c.drawRightString(w - 50, y, ar("خطاب توجيه تدريب تعاوني"))
    y -= 40

    c.setFont(FONT_NAME, 14)
    lines = [
        f"اسم المتدرب: {student['trainee_name']}",
        f"رقم المتدرب: {student['trainee_id']}",
        f"التخصص: {student['specialization']}",
        f"البرنامج: {student['program']}",
        f"جهة التدريب: {chosen['entity']}",
        "",
        "نأمل منكم التكرم باستقبال المتدرب المذكور أعلاه للتدريب التعاوني حسب الأنظمة المتبعة.",
        "شاكرين لكم تعاونكم،،",
    ]
    for line in lines:
        c.drawRightString(w - 50, y, ar(line))
        y -= 24

    c.setFont(FONT_NAME, 12)
    c.drawString(50, 40, ar(f"تم إنشاء الخطاب بتاريخ: {datetime.now().strftime('%Y-%m-%d')}"))

    c.showPage()
    c.save()

    return send_file(pdf_path, as_attachment=True, download_name=f"خطاب_توجيه_{tid}.pdf")


# Render port
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

import os
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
from flask import Flask, request, redirect, url_for, send_file, session, render_template_string, abort

from docx import Document


# =========================
# إعدادات عامة
# =========================
APP_TITLE = "بوابة خطاب التوجيه - التدريب التعاوني"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
STATIC_DIR = BASE_DIR / "static"
OUT_DIR = BASE_DIR / "out"

OUT_DIR.mkdir(exist_ok=True)

DATA_FILE = DATA_DIR / "students.xlsx"
ASSIGNMENTS_FILE = DATA_DIR / "assignments.json"

LETTER_TEMPLATE = DATA_DIR / "letter_template.docx"  # إذا موجود، سيتم استخدامه

SECRET_KEY = os.environ.get("SECRET_KEY", "change-me-please")

app = Flask(__name__)
app.secret_key = SECRET_KEY


# =========================
# أدوات مساعدة
# =========================
def _norm(s: str) -> str:
    """توحيد اسم العمود: إزالة المسافات + بعض الاختلافات البسيطة."""
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace("\u200f", "").replace("\u200e", "")
    s = s.replace("إ", "ا").replace("أ", "ا").replace("آ", "ا")
    s = s.replace("ة", "ه")  # تقليل اختلافات بسيطة
    s = " ".join(s.split())
    return s


def _digits_only(x) -> str:
    """تحويل القيم لأرقام كنص (بدون .0)"""
    if x is None:
        return ""
    s = str(x).strip()
    # معالجة أرقام الإكسل اللي تطلع 444229747.0
    if s.endswith(".0"):
        s = s[:-2]
    # إزالة أي شيء غير رقم
    out = "".join(ch for ch in s if ch.isdigit())
    return out


def load_students_df() -> pd.DataFrame:
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE)

    # تنظيف أسماء الأعمدة (مهم جدًا بسبب: 'جهة التدريب ' فيها مسافة)
    rename_map = {c: _norm(c) for c in df.columns}
    df = df.rename(columns=rename_map)

    # أسماء أعمدة متوقعة من ملفك
    # عندك: رقم المتدرب، رقم الجوال، التخصص، برنامج، جهة التدريب، اسم المقرر، الرقم المرجعي، المدرب، إسم المتدرب
    required = {
        "رقم المتدرب": ["رقم المتدرب", "رقم المتدرب"],
        "رقم الجوال": ["رقم الجوال", "رقم الجوال"],
        "التخصص": ["التخصص", "التخصص"],
        "برنامج": ["برنامج", "برنامج"],
        "جهة التدريب": ["جهه التدريب", "جهة التدريب", "جهة التدريب"],  # بعد _norm قد تصبح "جهه التدريب"
        "اسم المتدرب": ["اسم المتدرب", "اسم المتدرب", "اسم المتدرب", "اسم المتدرب"],
        "المدرب": ["المدرب", "المدرب"],
        "الرقم المرجعي": ["الرقم المرجعي", "الرقم المرجعي"],
        "اسم المقرر": ["اسم المقرر", "اسم المقرر"],
    }

    # إيجاد الأعمدة حتى لو اختلفت الكتابة قليلًا بعد التطبيع
    cols = {_norm(c): c for c in df.columns}

    def find_col(cands):
        for cand in cands:
            nc = _norm(cand)
            if nc in cols:
                return cols[nc]
        return None

    mapped = {}
    for key, cands in required.items():
        col = find_col(cands)
        if col is None:
            # بعض الأعمدة غير ضرورية 100% للتسجيل، لكن للخطاب نحتاجها.
            # سنسمح بغياب "المدرب/الرقم المرجعي/اسم المقرر" بدون كسر الدخول.
            if key in ["المدرب", "الرقم المرجعي", "اسم المقرر"]:
                mapped[key] = None
                continue
            raise ValueError(
                f"لم أجد عمود '{key}' داخل ملف الإكسل. الأعمدة الموجودة: {list(df.columns)}"
            )
        mapped[key] = col

    # إنشاء أعمدة موحدة بأسماء ثابتة نستخدمها في النظام
    out = pd.DataFrame()
    out["trainee_id"] = df[mapped["رقم المتدرب"]].apply(_digits_only)
    out["phone"] = df[mapped["رقم الجوال"]].apply(_digits_only)
    out["specialization"] = df[mapped["التخصص"]].astype(str).str.strip()
    out["program"] = df[mapped["برنامج"]].astype(str).str.strip()
    # بعد _norm قد يكون "جهه التدريب" أو "جهة التدريب"
    training_col = mapped["جهة التدريب"] or mapped.get("جهه التدريب")
    if training_col is None:
        # حاول إيجادها من df نفسه
        for c in df.columns:
            if _norm(c) in [_norm("جهة التدريب"), _norm("جهه التدريب")]:
                training_col = c
                break
    if training_col is None:
        raise ValueError("لم أجد عمود 'جهة التدريب' داخل ملف الإكسل.")
    out["entity"] = df[training_col].astype(str).str.strip()

    name_col = mapped["اسم المتدرب"]
    out["trainee_name"] = df[name_col].astype(str).str.strip()

    out["trainer"] = df[mapped["المدرب"]].astype(str).str.strip() if mapped["المدرب"] else ""
    out["ref_no"] = df[mapped["الرقم المرجعي"]].astype(str).str.strip() if mapped["الرقم المرجعي"] else ""
    out["course_name"] = df[mapped["اسم المقرر"]].astype(str).str.strip() if mapped["اسم المقرر"] else ""

    # تنظيف صفوف غير مكتملة
    out = out[(out["trainee_id"] != "") & (out["phone"] != "")]
    out = out.dropna(subset=["specialization", "entity"])
    out["entity"] = out["entity"].replace({"nan": ""})
    out = out[out["entity"].str.strip() != ""]

    return out


def load_assignments() -> dict:
    if not ASSIGNMENTS_FILE.exists():
        return {}
    try:
        return json.loads(ASSIGNMENTS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_assignments(data: dict) -> None:
    ASSIGNMENTS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def compute_base_slots(df: pd.DataFrame) -> dict:
    """
    الفرص = عدد تكرار الجهة داخل نفس التخصص في ملف الطلاب.
    الناتج:
      { specialization: { entity: count } }
    """
    base = {}
    grp = df.groupby(["specialization", "entity"]).size().reset_index(name="count")
    for _, row in grp.iterrows():
        spec = str(row["specialization"]).strip()
        ent = str(row["entity"]).strip()
        cnt = int(row["count"])
        base.setdefault(spec, {})[ent] = cnt
    return base


def compute_remaining_slots(df: pd.DataFrame, assignments: dict) -> dict:
    """
    remaining = base - assigned_count (ضمن نفس التخصص ونفس الجهة)
    """
    base = compute_base_slots(df)

    # عدّ الاختيارات
    used = {}
    for tid, info in assignments.items():
        spec = info.get("specialization", "").strip()
        ent = info.get("entity", "").strip()
        if not spec or not ent:
            continue
        used.setdefault(spec, {})
        used[spec][ent] = used[spec].get(ent, 0) + 1

    remaining = {}
    for spec, ents in base.items():
        remaining.setdefault(spec, {})
        for ent, cnt in ents.items():
            rem = cnt - used.get(spec, {}).get(ent, 0)
            remaining[spec][ent] = max(rem, 0)

    return remaining


def get_current_user():
    tid = session.get("tid")
    if not tid:
        return None
    return tid


# =========================
# واجهات HTML
# =========================
PAGE_LOGIN = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%; overflow:hidden; background:#fff;}
    .top-image img{width:100%; height:auto; display:block; object-fit:contain;}
    .wrap{max-width:1000px;margin:0 auto;padding:24px}
    .card{background:#fff;border-radius:18px;padding:28px;box-shadow:0 10px 25px rgba(0,0,0,.08)}
    h1{margin:0 0 10px;font-size:46px;text-align:center}
    .sub{margin:0 0 25px;text-align:center;color:#333;font-size:18px}
    .row{display:flex;gap:16px;flex-wrap:wrap;justify-content:space-between}
    .field{flex:1;min-width:280px}
    label{display:block;margin:10px 0 8px;font-weight:bold}
    input{width:100%;padding:14px;border-radius:14px;border:1px solid #ddd;font-size:20px}
    button{width:100%;padding:18px;border-radius:18px;border:0;background:#0b1730;color:#fff;font-size:24px;font-weight:bold;cursor:pointer;margin-top:18px}
    .err{color:#c00;font-weight:bold;margin-top:14px;text-align:center}
    .note{color:#666;margin-top:10px;text-align:center}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <p class="sub">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

      <form method="post" action="/">
        <div class="row">
          <div class="field">
            <label>الرقم التدريبي</label>
            <input name="tid" placeholder="مثال: 444229747" required>
          </div>
          <div class="field">
            <label>آخر 4 أرقام من الجوال</label>
            <input name="last4" placeholder="مثال: 5513" required>
          </div>
        </div>
        <button type="submit">دخول</button>
      </form>

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}

      <div class="note">ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b></div>
    </div>
  </div>
</body>
</html>
"""


PAGE_CHOOSE = r"""
<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="utf-8">
  <title>{{title}}</title>
  <style>
    body{font-family:Arial;background:#f7f7f7;margin:0}
    .top-image{width:100%; overflow:hidden; background:#fff;}
    .top-image img{width:100%; height:auto; display:block; object-fit:contain;}
    .wrap{max-width:1100px;margin:0 auto;padding:24px}
    .card{background:#fff;border-radius:18px;padding:28px;box-shadow:0 10px 25px rgba(0,0,0,.08)}
    h1{margin:0 0 10px;font-size:40px;text-align:center}
    .chips{display:flex;gap:12px;flex-wrap:wrap;justify-content:center;margin:18px 0}
    .chip{background:#eef3ff;border-radius:999px;padding:10px 14px;font-weight:bold}
    select{width:100%;padding:14px;border-radius:14px;border:1px solid #ddd;font-size:18px}
    button{width:100%;padding:18px;border-radius:18px;border:0;background:#0b1730;color:#fff;font-size:22px;font-weight:bold;cursor:pointer;margin-top:18px}
    .err{color:#c00;font-weight:bold;margin-top:14px;text-align:center}
    .ok{color:#0a7a2a;font-weight:bold;margin-top:14px;text-align:center}
    .box{background:#f7f9ff;border-radius:14px;padding:14px;margin-top:16px}
    a{color:#0645ad}
  </style>
</head>
<body>
  <div class="top-image">
    <img src="/static/header.jpg" alt="Header">
  </div>

  <div class="wrap">
    <div class="card">
      <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
      <div class="chips">
        <div class="chip">المتدرب: {{user.trainee_name}}</div>
        <div class="chip">رقم المتدرب: {{user.trainee_id}}</div>
        <div class="chip">التخصص/البرنامج: {{user.specialization}} — {{user.program}}</div>
      </div>

      {% if already %}
        <div class="box">
          <b>تم تسجيل اختيارك مسبقاً:</b><br>
          الجهة المختارة: <b>{{already}}</b><br><br>
          <a href="{{ url_for('letter', tid=user.trainee_id) }}">تحميل/طباعة خطاب التوجيه (Word)</a>
        </div>
      {% else %}
        <form method="post" action="{{ url_for('choose') }}">
          <label><b>جهة التدريب المتاحة (حسب تخصصك)</b></label>
          <select name="entity" required>
            <option value="">اختر الجهة...</option>
            {% for ent, rem in options %}
              <option value="{{ent}}">{{ent}} — (متبقي {{rem}})</option>
            {% endfor %}
          </select>
          <button type="submit">حفظ الاختيار</button>
        </form>
      {% endif %}

      {% if error %}
        <div class="err">{{error}}</div>
      {% endif %}
      {% if ok %}
        <div class="ok">{{ok}}</div>
      {% endif %}

      <div class="box">
        <b>ملخص الفرص المتبقية داخل تخصصك:</b>
        <ul>
          {% for ent, rem in summary %}
            <li>{{ent}} : {{rem}} فرصة</li>
          {% endfor %}
        </ul>
      </div>

      <div style="text-align:center;margin-top:10px">
        <a href="{{ url_for('logout') }}">تسجيل خروج</a>
      </div>

    </div>
  </div>
</body>
</html>
"""


# =========================
# Routes
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    # صفحة دخول
    if request.method == "GET":
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=None)

    # POST (تحقق)
    try:
        df = load_students_df()
    except Exception as e:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error=f"خطأ أثناء قراءة ملف الطلاب: {e}")

    tid = _digits_only(request.form.get("tid"))
    last4 = _digits_only(request.form.get("last4"))

    if not tid or not last4:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="يرجى إدخال البيانات بشكل صحيح.")

    row = df[df["trainee_id"] == tid]
    if row.empty:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="رقم المتدرب غير موجود في الملف.")

    phone = row.iloc[0]["phone"]
    if len(phone) < 4 or phone[-4:] != last4:
        return render_template_string(PAGE_LOGIN, title=APP_TITLE, error="بيانات الدخول غير صحيحة. تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

    session["tid"] = tid
    return redirect(url_for("choose"))


@app.route("/choose", methods=["GET", "POST"])
def choose():
    tid = get_current_user()
    if not tid:
        return redirect(url_for("index"))

    df = load_students_df()
    user_row = df[df["trainee_id"] == tid].iloc[0].to_dict()

    assignments = load_assignments()
    remaining = compute_remaining_slots(df, assignments)

    spec = user_row["specialization"]
    spec_remaining = remaining.get(spec, {})

    # خيارات تظهر فقط للجهات الموجودة في نفس تخصص المتدرب + المتبقي > 0
    options = [(ent, rem) for ent, rem in spec_remaining.items() if rem > 0]
    options.sort(key=lambda x: (-x[1], x[0]))

    already = None
    if tid in assignments:
        already = assignments[tid].get("entity")

    if request.method == "POST":
        if already:
            return redirect(url_for("choose"))

        entity = (request.form.get("entity") or "").strip()
        if entity not in spec_remaining:
            return render_template_string(
                PAGE_CHOOSE,
                title=APP_TITLE,
                user=user_row,
                options=options,
                summary=sorted(spec_remaining.items(), key=lambda x: (-x[1], x[0])),
                already=None,
                error="الجهة المختارة غير صحيحة.",
                ok=None,
            )

        if spec_remaining.get(entity, 0) <= 0:
            return render_template_string(
                PAGE_CHOOSE,
                title=APP_TITLE,
                user=user_row,
                options=options,
                summary=sorted(spec_remaining.items(), key=lambda x: (-x[1], x[0])),
                already=None,
                error="هذه الجهة لم تعد متاحة (انتهت الفرص). اختر جهة أخرى.",
                ok=None,
            )

        # حفظ الاختيار
        assignments[tid] = {
            "trainee_id": tid,
            "trainee_name": user_row.get("trainee_name", ""),
            "phone": user_row.get("phone", ""),
            "specialization": spec,
            "program": user_row.get("program", ""),
            "entity": entity,
            "saved_at": datetime.utcnow().isoformat(),
            "trainer": user_row.get("trainer", ""),
            "ref_no": user_row.get("ref_no", ""),
            "course_name": user_row.get("course_name", ""),
        }
        save_assignments(assignments)

        return redirect(url_for("choose"))

    # GET
    summary = sorted(spec_remaining.items(), key=lambda x: (-x[1], x[0]))
    return render_template_string(
        PAGE_CHOOSE,
        title=APP_TITLE,
        user=user_row,
        options=options,
        summary=summary,
        already=already,
        error=None,
        ok=None,
    )


@app.route("/letter")
def letter():
    tid = request.args.get("tid", "")
    tid = _digits_only(tid)
    if not tid:
        abort(404)

    assignments = load_assignments()
    info = assignments.get(tid)
    if not info:
        return "لا يوجد اختيار محفوظ لهذا المتدرب.", 400

    # توليد ملف Word
    OUT_DIR.mkdir(exist_ok=True)
    out_path = OUT_DIR / f"خطاب_توجيه_{tid}.docx"

    if LETTER_TEMPLATE.exists():
        doc = Document(str(LETTER_TEMPLATE))
    else:
        # قالب بسيط احتياطي
        doc = Document()
        doc.add_paragraph("خطاب توجيه متدرب تدريب تعاوني")
        doc.add_paragraph("")

    # محاولة تعبئة القالب (placeholders) + تعبئة جدول لو موجود
    def replace_in_paragraph(paragraph, mapping):
        for key, val in mapping.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key, val)

    mapping = {
        "{{NAME}}": info.get("trainee_name", ""),
        "{{TRAINEE_ID}}": info.get("trainee_id", ""),
        "{{PHONE}}": info.get("phone", ""),
        "{{SPECIALIZATION}}": info.get("specialization", ""),
        "{{PROGRAM}}": info.get("program", ""),
        "{{ENTITY}}": info.get("entity", ""),
        "{{REF_NO}}": info.get("ref_no", ""),
        "{{TRAINER}}": info.get("trainer", ""),
        "{{COURSE_NAME}}": info.get("course_name", ""),
    }

    # Replace in paragraphs
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)

    # Replace in tables and also try to fill known table positions
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p, mapping)

    # إذا القالب فيه الجدول مثل نموذجك (الصف الأول بيانات المتدرب)
    # نحاول تعبئة أول جدول (لو موجود) بطريقة ذكية
    if doc.tables:
        t = doc.tables[0]
        # نتحقق أن الجدول على الأقل 2 صفوف
        if len(t.rows) >= 2 and len(t.rows[1].cells) >= 5:
            # الصف الثاني غالباً هو صف البيانات تحت العناوين
            # ترتيب الأعمدة من صورتك: الرقم | الاسم | الرقم الأكاديمي | التخصص | جوال
            # سنحاول تعبئته بدون تخريب
            try:
                cells = t.rows[1].cells
                # بعض القوالب بالعكس RTL - لذلك نملأ حسب عدد الخلايا
                # الأفضل: نبحث عن عناوين في الصف الأول
                headers = [c.text.strip() for c in t.rows[0].cells]
                # fallback: تعبئة منطقيّة
                if any("الرقم" in h for h in headers) and any("الاسم" in h for h in headers):
                    # اتركها؛ لأن القالب قد مضبوط بالـ placeholders أصلاً
                    pass
                else:
                    cells[-1].text = info.get("trainee_id", "")          # الرقم
                    cells[-2].text = info.get("trainee_name", "")        # الاسم
                    cells[-3].text = info.get("trainee_id", "")          # الرقم الأكاديمي (نفس الرقم لو ما فيه غيره)
                    cells[-4].text = info.get("specialization", "")      # التخصص
                    cells[-5].text = info.get("phone", "")               # جوال
            except Exception:
                pass

    doc.save(str(out_path))

    return send_file(
        str(out_path),
        as_attachment=True,
        download_name=out_path.name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("index"))


# =========================
# تشغيل
# =========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

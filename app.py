from __future__ import annotations

import json
import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, request, redirect, url_for, render_template_string, send_file

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "students.xlsx"
ASSIGNMENTS_FILE = BASE_DIR / "data" / "assignments.json"


# -----------------------------
# Helpers
# -----------------------------
def norm(s: str) -> str:
    """Normalize column names (remove extra spaces, unify)"""
    return " ".join(str(s).strip().split())


def load_students_df() -> pd.DataFrame:
    if not DATA_FILE.exists():
        raise FileNotFoundError(f"الملف غير موجود: {DATA_FILE}")

    df = pd.read_excel(DATA_FILE)

    # Normalize columns (important: handles 'جهة التدريب ' with trailing space)
    df.columns = [norm(c) for c in df.columns]

    # Expected Arabic columns (your sheet)
    # We will auto-detect even if slight variations exist.
    possible = {
        "trainee_id": ["رقم المتدرب", "الرقم التدريبي", "رقم تدريبي", "رقم_المتدرب"],
        "phone": ["رقم الجوال", "الجوال", "رقم الهاتف", "هاتف", "موبايل"],
        "trainee_name": ["اسم المتدرب", "إسم المتدرب", "اسم_المتدرب"],
        "program": ["برنامج", "البرنامج"],
        "major": ["التخصص", "تخصص"],
        "entity": ["جهة التدريب", "جهة التدريب", "الجهة التدريبية", "جهة_التدريب"],
        "trainer": ["المدرب", "مدرب"],
        "ref_no": ["الرقم المرجعي", "رقم مرجعي", "مرجعي"],
        "course_name": ["اسم المقرر", "المقرر", "اسم مقرر"],
    }

    def pick_col(keys: list[str]) -> str | None:
        cols = set(df.columns)
        for k in keys:
            k2 = norm(k)
            if k2 in cols:
                return k2
        # extra fallback: contains match
        for c in df.columns:
            for k in keys:
                if norm(k) in c:
                    return c
        return None

    col_map = {k: pick_col(v) for k, v in possible.items()}

    # Required columns for login + opportunities
    required = ["trainee_id", "phone", "trainee_name", "major", "program", "entity"]
    missing = [r for r in required if not col_map.get(r)]
    if missing:
        raise KeyError(
            f"لم يتم العثور على الأعمدة المطلوبة: {missing}. "
            f"الأعمدة الموجودة: {list(df.columns)}"
        )

    # Keep only needed cols (but keep others if you want later)
    df = df.rename(columns={
        col_map["trainee_id"]: "trainee_id",
        col_map["phone"]: "phone",
        col_map["trainee_name"]: "trainee_name",
        col_map["major"]: "major",
        col_map["program"]: "program",
        col_map["entity"]: "entity",
        (col_map["trainer"] or "trainer"): "trainer",
        (col_map["ref_no"] or "ref_no"): "ref_no",
        (col_map["course_name"] or "course_name"): "course_name",
    })

    # Clean values
    df["trainee_id"] = df["trainee_id"].astype(str).str.strip()
    df["phone"] = df["phone"].astype(str).str.strip()
    df["trainee_name"] = df["trainee_name"].astype(str).str.strip()
    df["major"] = df["major"].astype(str).str.strip()
    df["program"] = df["program"].astype(str).str.strip()
    df["entity"] = df["entity"].astype(str).str.strip()

    return df


def load_assignments() -> dict:
    if not ASSIGNMENTS_FILE.exists():
        return {}
    try:
        return json.loads(ASSIGNMENTS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_assignments(data: dict) -> None:
    ASSIGNMENTS_FILE.parent.mkdir(parents=True, exist_ok=True)
    ASSIGNMENTS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def last4(phone: str) -> str:
    digits = "".join([ch for ch in str(phone) if ch.isdigit()])
    return digits[-4:] if len(digits) >= 4 else digits


def get_slots_by_major(df: pd.DataFrame) -> dict[str, dict[str, int]]:
    """
    Slots are computed ONLY from Excel:
    For each major: count repeated training entities => available slots.
    """
    slots: dict[str, dict[str, int]] = {}
    tmp = df.copy()
    tmp["major"] = tmp["major"].astype(str).str.strip()
    tmp["entity"] = tmp["entity"].astype(str).str.strip()

    tmp = tmp[(tmp["major"] != "") & (tmp["entity"] != "") & (tmp["entity"].str.lower() != "nan")]

    grouped = tmp.groupby(["major", "entity"]).size().reset_index(name="count")
    for _, row in grouped.iterrows():
        m = str(row["major"]).strip()
        e = str(row["entity"]).strip()
        c = int(row["count"])
        slots.setdefault(m, {})[e] = c
    return slots


def remaining_slots_for_major(df: pd.DataFrame, major: str) -> dict[str, int]:
    slots_all = get_slots_by_major(df)
    total = slots_all.get(major, {}).copy()

    assignments = load_assignments()
    # subtract chosen assignments for same major+entity
    for tid, info in assignments.items():
        if not isinstance(info, dict):
            continue
        if str(info.get("major", "")).strip() != major:
            continue
        ent = str(info.get("entity", "")).strip()
        if ent in total:
            total[ent] = max(0, total[ent] - 1)

    # keep only > 0
    return {k: v for k, v in total.items() if v > 0}


# -----------------------------
# Pages (templates inline)
# -----------------------------
BASE_CSS = """
<style>
  body{font-family:Tahoma,Arial; background:#f6f6f6; margin:0; direction:rtl;}
  .top-image{background:#fff; padding:10px 0; text-align:center;}
  .top-image img{max-width:100%; height:auto; max-height:220px; object-fit:contain;}
  .wrap{max-width:1050px; margin:20px auto; padding:0 14px;}
  .card{background:#fff; border-radius:18px; padding:28px; box-shadow:0 6px 20px rgba(0,0,0,.06);}
  h1{margin:0 0 10px; font-size:44px; text-align:center;}
  .muted{color:#666; text-align:center;}
  .row{display:flex; gap:18px; margin-top:18px; flex-wrap:wrap;}
  .col{flex:1; min-width:260px;}
  label{display:block; font-weight:bold; margin-bottom:8px;}
  input, select{width:100%; padding:14px; border-radius:14px; border:1px solid #ddd; font-size:18px;}
  button{width:100%; padding:16px; border-radius:16px; border:0; background:#07152f; color:#fff; font-size:20px; cursor:pointer;}
  button:hover{opacity:.95;}
  .err{color:#b00020; margin-top:12px; font-weight:bold; text-align:center;}
  .ok{color:#0a7a2f; margin-top:12px; font-weight:bold; text-align:center;}
  .pillrow{display:flex; gap:10px; justify-content:center; flex-wrap:wrap; margin:10px 0 22px;}
  .pill{background:#eef4ff; padding:10px 14px; border-radius:999px; font-weight:bold;}
  .small{font-size:14px; color:#666; text-align:center; margin-top:8px;}
  a{color:#0a3cff; text-decoration:underline;}
</style>
"""

LOGIN_HTML = """
<!doctype html>
<html lang="ar"><head><meta charset="utf-8"><title>بوابة خطاب التوجيه</title>
""" + BASE_CSS + """
</head>
<body>
<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>
<div class="wrap">
  <div class="card">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <p class="muted">يرجى تسجيل الدخول بالرقم التدريبي وآخر 4 أرقام من رقم الجوال.</p>

    <form method="post" action="/">
      <div class="row">
        <div class="col">
          <label>الرقم التدريبي</label>
          <input name="trainee_id" placeholder="مثال: 444229747" required>
        </div>
        <div class="col">
          <label>آخر 4 أرقام من الجوال</label>
          <input name="last4" placeholder="مثال: 6101" required>
        </div>
      </div>
      <div style="margin-top:18px;">
        <button type="submit">دخول</button>
      </div>
    </form>

    {% if error %}
      <div class="err">{{error}}</div>
    {% endif %}

    <div class="small">ملاحظة: يتم قراءة ملف الطلاب من <b>data/students.xlsx</b></div>
  </div>
</div>
</body></html>
"""

CHOOSE_HTML = """
<!doctype html>
<html lang="ar"><head><meta charset="utf-8"><title>اختيار جهة التدريب</title>
""" + BASE_CSS + """
</head>
<body>
<div class="top-image">
  <img src="/static/header.jpg" alt="Header">
</div>

<div class="wrap">
  <div class="card">
    <h1>بوابة خطاب التوجيه - التدريب التعاوني</h1>
    <p class="muted">اختر جهة التدريب المتاحة لتخصصك ثم احفظ الاختيار.</p>

    <div class="pillrow">
      <div class="pill">المتدرب: {{name}}</div>
      <div class="pill">رقم المتدرب: {{tid}}</div>
      <div class="pill">التخصص/البرنامج: {{major}} — {{program}}</div>
    </div>

    {% if already %}
      <div class="ok">
        تم تسجيل اختيارك مسبقًا:<br>
        <b>الجهة المختارة: {{already}}</b><br><br>
        <a href="/letter?tid={{tid}}">تحميل/طباعة خطاب التوجيه PDF</a>
      </div>
      <hr style="margin:22px 0;">
    {% endif %}

    <form method="post" action="/choose?tid={{tid}}">
      <div class="row">
        <div class="col">
          <label>جهة التدريب المتاحة</label>
          <select name="entity" required>
            <option value="">اختر الجهة...</option>
            {% for e, c in options %}
              <option value="{{e}}">{{e}} ({{c}} فرص متبقية)</option>
            {% endfor %}
          </select>
        </div>
      </div>

      <div style="margin-top:18px;">
        <button type="submit">حفظ الاختيار</button>
      </div>
    </form>

    {% if error %}
      <div class="err">{{error}}</div>
    {% endif %}

    {% if summary %}
      <hr style="margin:22px 0;">
      <h3 style="margin:0 0 10px;">ملخص الفرص المتبقية حسب الجهات داخل تخصصك:</h3>
      <ul>
        {% for e, c in summary %}
          <li>{{e}} : <b>{{c}}</b> فرصة</li>
        {% endfor %}
      </ul>
    {% endif %}

  </div>
</div>
</body></html>
"""


# -----------------------------
# Routes
# -----------------------------
@app.get("/login")
def login_redirect():
    return redirect(url_for("login_page"))


@app.get("/")
def login_page():
    return render_template_string(LOGIN_HTML, error=None)


@app.post("/")
def login_post():
    try:
        df = load_students_df()
    except Exception as e:
        return render_template_string(LOGIN_HTML, error=f"خطأ أثناء التحميل/التحقق: {e}")

    tid = (request.form.get("trainee_id") or "").strip()
    l4 = (request.form.get("last4") or "").strip()

    if not tid or not l4:
        return render_template_string(LOGIN_HTML, error="يرجى إدخال الرقم التدريبي وآخر 4 أرقام من الجوال.")

    row = df[df["trainee_id"] == tid]
    if row.empty:
        return render_template_string(LOGIN_HTML, error="الرقم التدريبي غير موجود.")

    phone = str(row.iloc[0]["phone"]).strip()
    if last4(phone) != last4(l4):
        return render_template_string(LOGIN_HTML, error="بيانات الدخول غير صحيحة، تأكد من الرقم التدريبي وآخر 4 أرقام من الجوال.")

    return redirect(url_for("choose_page", tid=tid))


@app.get("/choose")
def choose_page():
    tid = (request.args.get("tid") or "").strip()
    if not tid:
        return redirect(url_for("login_page"))

    try:
        df = load_students_df()
    except Exception as e:
        return f"خطأ: {e}", 500

    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("login_page"))

    name = str(row.iloc[0]["trainee_name"]).strip()
    major = str(row.iloc[0]["major"]).strip()
    program = str(row.iloc[0]["program"]).strip()

    # remaining slots for this major
    remaining = remaining_slots_for_major(df, major)
    options = sorted(remaining.items(), key=lambda x: (-x[1], x[0]))
    summary = options

    assignments = load_assignments()
    already = None
    if tid in assignments and isinstance(assignments[tid], dict):
        already = assignments[tid].get("entity")

    return render_template_string(
        CHOOSE_HTML,
        tid=tid,
        name=name,
        major=major,
        program=program,
        options=options,
        summary=summary,
        already=already,
        error=None
    )


@app.post("/choose")
def choose_post():
    tid = (request.args.get("tid") or "").strip()
    entity = (request.form.get("entity") or "").strip()
    if not tid:
        return redirect(url_for("login_page"))
    if not entity:
        return redirect(url_for("choose_page", tid=tid))

    df = load_students_df()
    row = df[df["trainee_id"] == tid]
    if row.empty:
        return redirect(url_for("login_page"))

    major = str(row.iloc[0]["major"]).strip()

    remaining = remaining_slots_for_major(df, major)
    if remaining.get(entity, 0) <= 0:
        # reload choose with error
        name = str(row.iloc[0]["trainee_name"]).strip()
        program = str(row.iloc[0]["program"]).strip()
        options = sorted(remaining.items(), key=lambda x: (-x[1], x[0]))
        return render_template_string(
            CHOOSE_HTML,
            tid=tid,
            name=name,
            major=major,
            program=program,
            options=options,
            summary=options,
            already=None,
            error="عذرًا، هذه الجهة لم تعد متاحة (انتهت الفرص). اختر جهة أخرى."
        )

    assignments = load_assignments()
    assignments[tid] = {
        "entity": entity,
        "major": major,
        "saved_at": datetime.utcnow().isoformat() + "Z"
    }
    save_assignments(assignments)

    return redirect(url_for("choose_page", tid=tid))


@app.get("/letter")
def letter_pdf():
    # ملاحظة: هذا المسار تركته موجودًا حتى لا يعطي Not Found
    # حالياً يرجّع رسالة بسيطة بدل خطأ.
    tid = (request.args.get("tid") or "").strip()
    if not tid:
        return redirect(url_for("login_page"))
    return (
        "تم حفظ الاختيار بنجاح. خطوة PDF النهائية سنثبتها بعد اعتماد القالب النهائي للطباعة.",
        200
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port)

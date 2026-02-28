import json
import pandas as pd
from pathlib import Path

# ====== إعدادات المسارات ======
STUDENTS_XLSX = Path("data/students.xlsx")
OUTPUT_JSON = Path("data/slots.json")

# ====== قراءة ملف الطلاب ======
if not STUDENTS_XLSX.exists():
    raise FileNotFoundError(f"لم أجد الملف: {STUDENTS_XLSX.resolve()}")

df = pd.read_excel(STUDENTS_XLSX)

# ====== تحديد اسم عمود التخصص (قد يختلف) ======
possible_cols = ["التخصص", "التخصص ", "Specialty", "specialty", "تخصص"]
spec_col = None
for c in df.columns:
    if str(c).strip() in [x.strip() for x in possible_cols] or str(c).strip() == "التخصص":
        spec_col = c
        break

if spec_col is None:
    # اطبع الأعمدة لمساعدتك
    raise ValueError(f"لم أجد عمود التخصص. الأعمدة الموجودة: {list(df.columns)}")

# ====== تنظيف التخصصات وحساب العدد ======
spec_series = df[spec_col].astype(str).str.strip()
spec_series = spec_series[spec_series.notna() & (spec_series != "") & (spec_series.str.lower() != "nan")]

counts = spec_series.value_counts().sort_index()

# ====== بناء slots.json (كل تخصص فرصة واحدة بمقاعد = عدد الطلاب) ======
slots = []
for specialty, cnt in counts.items():
    slots.append({
        "program": specialty,                 # البرنامج = اسم التخصص
        "training_entity": "حسب التخصص",      # ثابت مؤقت (يُستخدم لو النظام يحتاجه)
        "job_title": f"فرصة تدريبية - {specialty}",
        "capacity": int(cnt)
    })

# ====== حفظ الناتج ======
OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
    json.dump(slots, f, ensure_ascii=False, indent=2)

print("✅ تم إنشاء الملف بنجاح:", OUTPUT_JSON.resolve())
print("عدد التخصصات (الفرص):", len(slots))
print("مثال أول 3 فرص:\n", json.dumps(slots[:3], ensure_ascii=False, indent=2))

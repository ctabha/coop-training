import os
import pandas as pd
from flask import Flask, render_template, request

app = Flask(__name__)

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
STUDENTS_XLSX = os.path.join(DATA_DIR, "students.xlsx")
SLOTS_JSON = os.path.join(DATA_DIR, "slots.json")
ASSIGNMENTS_JSON = os.path.join(DATA_DIR, "assignments.json")

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    # توحيد أسماء الأعمدة (إزالة مسافات/أسطر/رموز غريبة)
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace("\r", " ", regex=False)
        .str.strip()
    )
    return df

def load_students_df() -> pd.DataFrame:
    df = pd.read_excel(STUDENTS_XLSX, dtype=str).fillna("")
    df = normalize_cols(df)
    # نظّف أرقام الهوية والجوال من الفراغات
    if "رقم المتدرب" in df.columns:
        df["رقم المتدرب"] = df["رقم المتدرب"].str.strip()
    if "رقم الجوال" in df.columns:
        df["رقم الجوال"] = df["رقم الجوال"].str.strip()
    return df

def find_trainee(df: pd.DataFrame, trainee_id: str, last4: str) -> dict | None:
    trainee_id = (trainee_id or "").strip()
    last4 = (last4 or "").strip()

    # تأكد من وجود الأعمدة الأساسية
    required = {"رقم المتدرب", "رقم الجوال"}
    if not required.issubset(set(df.columns)):
        # رجّع None وخلي الرسالة واضحة للمستخدم بدل Exception
        return None

    # فلترة: رقم المتدرب + آخر 4 من الجوال
    m = df[df["رقم المتدرب"].eq(trainee_id)]
    if m.empty:
        return None

    # آخر 4 أرقام من الجوال
    def phone_last4(x: str) -> str:
        x = "".join([c for c in str(x) if c.isdigit()])
        return x[-4:] if len(x) >= 4 else ""

    m = m[m["رقم الجوال"].apply(phone_last4).eq(last4)]
    if m.empty:
        return None

    # ✅ الأهم: رجّع قاموس واحد وليس قائمة
    return m.iloc[0].to_dict()

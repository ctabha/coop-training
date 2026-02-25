import os
import io
import pandas as pd
from flask import Flask, request, send_file, abort

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import arabic_reshaper
from bidi.algorithm import get_display

from pypdf import PdfReader, PdfWriter

app = Flask(__name__)

DATA_FILE = os.path.join("data", "students.xlsx")
TEMPLATE_PDF = os.path.join("data", "letter_template.pdf")
FONT_PATH = os.path.join("static", "fonts", "NotoNaskhArabic-Regular.ttf")

pdfmetrics.registerFont(TTFont("AR", FONT_PATH))


def rtl(text):
    if text is None:
        return ""
    reshaped = arabic_reshaper.reshape(str(text))
    return get_display(reshaped)


def load_students():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.str.strip()
    return df


@app.get("/letter")
def letter():
    tid = request.args.get("tid", "").strip()
    if not tid:
        abort(400)

    df = load_students()

    required = [
        "رقم المتدرب",
        "اسم المتدرب",
        "رقم الجوال",
        "التخصص",
        "المدرب",
        "الرقم المرجعي",
        "جهة التدريب"
    ]

    missing = [c for c in required if c not in df.columns]
    if missing:
        return f"الأعمدة ناقصة: {missing}", 500

    row = df.loc[df["رقم المتدرب"].astype(str).str.strip() == tid]

    if row.empty:
        return "لم يتم العثور على المتدرب.", 404

    row = row.iloc[0]

    if not row["جهة التدريب"]:
        return "لم يتم العثور على اختيار محفوظ لهذا المتدرب.", 404

    trainee_id = row["رقم المتدرب"]
    name = row["اسم المتدرب"]
    phone = row["رقم الجوال"]
    major = row["التخصص"]
    supervisor = row["المدرب"]
    ref_no = row["الرقم المرجعي"]
    entity = row["جهة التدريب"]

    base_pdf = PdfReader(TEMPLATE_PDF)
    first_page = base_pdf.pages[0]

    width = float(first_page.mediabox.width)
    height = float(first_page.mediabox.height)

    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("AR", 14)

    # عدل الإحداثيات فقط إذا احتجت ضبط مكان النص
    c.drawRightString(width - 60, height - 260, rtl(str(trainee_id)))
    c.drawRightString(width - 60, height - 300, rtl(name))
    c.drawRightString(width - 60, height - 340, rtl(str(trainee_id)))
    c.drawRightString(width - 60, height - 380, rtl(major))
    c.drawRightString(width - 60, height - 420, rtl(str(phone)))
    c.drawRightString(width - 60, height - 460, rtl(supervisor))
    c.drawRightString(width - 60, height - 500, rtl(str(ref_no)))
    c.drawRightString(width - 60, height - 540, rtl(entity))

    c.save()
    packet.seek(0)

    overlay_pdf = PdfReader(packet)
    overlay_page = overlay_pdf.pages[0]

    merged_page = base_pdf.pages[0]
    merged_page.merge_page(overlay_page)

    writer = PdfWriter()
    writer.add_page(merged_page)

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"خطاب_توجيه_{trainee_id}.pdf",
        mimetype="application/pdf"
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)

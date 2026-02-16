# بوابة خطاب التوجيه - التدريب التعاوني (Python)

## التشغيل محليًا
```bash
pip install -r requirements.txt
python app.py
```
ثم افتح:
- http://127.0.0.1:8000

## الملفات
- data/trainees.xlsx  ملف المتدربين (مرفق)
- data/letter_template.docx قالب الخطاب (مرفق)
- data/slots.json عدّاد الفرص لكل جهة
- data/assignments.json حفظ حجوزات المتدربين (حتى لا يُخصم مرتين)

## ملاحظة
توليد PDF يتم عبر LibreOffice (soffice) على الخادم.

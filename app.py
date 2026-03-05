import os
import json
import pandas as pd
from flask import Flask,render_template,request,redirect,url_for,send_file
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics

BASE_DIR=os.path.dirname(os.path.abspath(__file__))

DATA=os.path.join(BASE_DIR,"data")
STATIC=os.path.join(BASE_DIR,"static")

STUDENTS=os.path.join(DATA,"students.xlsx")
ASSIGN=os.path.join(DATA,"assignments.json")
TEMPLATE=os.path.join(DATA,"letter_template.docx")

FONT=os.path.join(STATIC,"fonts","NotoNaskhArabic-Regular.ttf")

app=Flask(__name__)


def load_json(path):

    if not os.path.exists(path):
        return {}

    try:
        with open(path,"r",encoding="utf8") as f:
            return json.load(f)
    except:
        return {}


def save_json(path,data):

    with open(path,"w",encoding="utf8") as f:
        json.dump(data,f,ensure_ascii=False,indent=2)



def load_students():

    df=pd.read_excel(STUDENTS)

    students=[]

    for _,r in df.iterrows():

        students.append({

        "trainee_id":str(r[0]).strip(),
        "phone":str(r[1]).strip(),
        "name":str(r[2]).strip(),
        "specialty":str(r[3]).strip(),
        "program":str(r[4]).strip(),
        "entity":str(r[5]).strip()

        })

    return students



def build_entities():

    students=load_students()

    entities={}

    for s in students:

        spec=s["specialty"]
        ent=s["entity"]

        if spec not in entities:
            entities[spec]=[]

        if ent not in entities[spec]:
            entities[spec].append(ent)

    return entities



@app.route("/",methods=["GET","POST"])
def index():

    if request.method=="GET":
        return render_template("index.html",error=None)

    trainee_id=request.form.get("trainee_id")
    phone=request.form.get("phone_last4")

    students=load_students()

    for s in students:

        if s["trainee_id"]==trainee_id and s["phone"][-4:]==phone:

            return redirect(url_for("dashboard",trainee_id=trainee_id))

    return render_template("index.html",error="بيانات الدخول غير صحيحة")



@app.route("/dashboard/<trainee_id>")
def dashboard(trainee_id):

    students=load_students()

    trainee=None

    for s in students:
        if s["trainee_id"]==trainee_id:
            trainee=s

    assignments=load_json(ASSIGN)

    chosen=None

    if trainee_id in assignments:
        chosen=assignments[trainee_id]["entity"]

    entities=build_entities().get(trainee["specialty"],[])

    return render_template("dashboard.html",
                           trainee=trainee,
                           chosen=chosen,
                           entities=entities)



@app.route("/choose/<trainee_id>",methods=["POST"])
def choose(trainee_id):

    entity=request.form.get("entity")

    assignments=load_json(ASSIGN)

    assignments[trainee_id]={

    "entity":entity

    }

    save_json(ASSIGN,assignments)

    return redirect(url_for("dashboard",trainee_id=trainee_id))



@app.route("/pdf/<trainee_id>")
def pdf(trainee_id):

    assignments=load_json(ASSIGN)

    students=load_students()

    trainee=None

    for s in students:
        if s["trainee_id"]==trainee_id:
            trainee=s

    entity=assignments[trainee_id]["entity"]

    mapping={

    "{{trainee_name}}":trainee["name"],
    "{{trainee_id}}":trainee["trainee_id"],
    "{{phone}}":trainee["phone"],
    "{{training_entity}}":entity

    }

    doc=Document(TEMPLATE)

    for p in doc.paragraphs:
        for k,v in mapping.items():
            if k in p.text:
                p.text=p.text.replace(k,v)

    temp=os.path.join(BASE_DIR,"letter.docx")
    doc.save(temp)

    pdfmetrics.registerFont(TTFont("arabic",FONT))

    pdf_path=os.path.join(BASE_DIR,"letter.pdf")

    c=canvas.Canvas(pdf_path,pagesize=A4)

    y=800

    for p in doc.paragraphs:

        text=p.text.strip()

        if text:

            c.setFont("arabic",14)

            c.drawRightString(550,y,text)

            y-=20

    c.save()

    return send_file(pdf_path,as_attachment=True)

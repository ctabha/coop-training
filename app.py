from flask import send_file, abort
@app.route("/letter")
def letter_pdf():
    tid = str(request.args.get("tid", "")).strip()
    if not tid:
        abort(404)

    df = load_students_df()
    ensure_required_columns(df)

    row = df[df[COL_TRAINEE_ID].astype(str) == tid]
    if row.empty:
        abort(404)

    assignments = get_assignments()
    chosen_entity = assignments.get(tid)
    if not chosen_entity:
        return "لم يتم تسجيل جهة تدريب لهذا المتدرب بعد.", 400

    r = row.iloc[0].to_dict()

    student = {
        "رقم المتدرب": str(r.get(COL_TRAINEE_ID, "")).strip(),
        "إسم المتدرب": str(r.get(COL_TRAINEE_NAME, "")).strip(),
        "رقم الجوال": str(r.get(COL_PHONE, "")).strip(),
        "التخصص": str(r.get(COL_SPECIALIZATION, "")).strip(),
        "برنامج": str(r.get(COL_PROGRAM, "")).strip(),
        "المدرب": str(r.get(COL_TRAINER, "")).strip(),
        "الرقم المرجعي": str(r.get(COL_REFNO, "")).strip(),
    }

    html = build_letter_html(student, chosen_entity)

    pdf_bytes = HTML(string=html, base_url=str(BASE_DIR)).write_pdf()

    out_path = DATA_DIR / f"letter_{tid}.pdf"
    out_path.write_bytes(pdf_bytes)

    return send_file(
        out_path,
        mimetype="application/pdf",
        as_attachment=True,          # مهم: يجبره تحميل PDF
        download_name=f"خطاب_توجيه_{tid}.pdf",
    )

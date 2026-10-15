from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import io

app = Flask(__name__)

# १. मुख्य होमपेज
@app.route('/')
def home():
    return render_template('home.html')

# २. दस्त मसुदा मेनू
@app.route('/draft')
def draft():
    return render_template('draft_menu.html')

# ३. खरेदीखत उप-मेनू
@app.route('/kharedi_menu')
def kharedi_menu():
    return render_template('kharedi_menu.html')

# ४. एकसमान युनिव्हर्सल फॉर्म (सर्व ड्राफ्ट्ससाठी)
@app.route('/universal_draft')
@app.route('/plot_draft')
@app.route('/karar_to_kharedi_draft')
def universal_draft():
    return render_template('universal_draft_form.html')

# ५. मेनूमधून राउटिंग
@app.route('/draft_form')
def draft_form():
    return render_template('universal_draft_form.html')

# ६. सर्व ड्राफ्ट्सची अचूक Word (.docx) निर्मिती
@app.route('/generate_universal_docx', methods=['POST'])
@app.route('/generate_plot_docx', methods=['POST'])
@app.route('/generate_karar_to_kharedi_docx', methods=['POST'])
def generate_universal_docx():
    f = request.form
    doc_type = f.get('doc_type', 'प्लॉट खरेदीपत्र')
    doc = Document()

    # पक्षकार वर्गीकरण
    roles = f.getlist('party_role[]')
    names = f.getlist('party_name[]')
    ages = f.getlist('party_age[]')
    occs = f.getlist('party_occ[]')
    addrs = f.getlist('party_address[]')
    aadhars = f.getlist('party_aadhar[]')
    pans = f.getlist('party_pan[]')
    mobs = f.getlist('party_mobile[]')

    ghenar_list, denar_list, poa_list, manyata_list, jamin_list, sakshi_list = [], [], [], [], [], []

    for i in range(len(names)):
        item = {
            "name": names[i],
            "age": ages[i] if i < len(ages) else "",
            "occ": occs[i] if i < len(occs) else "",
            "addr": addrs[i] if i < len(addrs) else "",
            "aadhar": aadhars[i] if i < len(aadhars) else "",
            "pan": pans[i] if i < len(pans) else "",
            "mob": mobs[i] if i < len(mobs) else ""
        }
        role = roles[i] if i < len(roles) else "लिहून देणार"

        if role == 'लिहून घेणार':
            ghenar_list.append(item)
        elif role == 'लिहून देणार':
            denar_list.append(item)
        elif 'वटमुखत्यार' in role:
            poa_list.append(item)
        elif 'मान्यता' in role or 'संमती' in role:
            manyata_list.append(item)
        elif 'जामीनदार' in role:
            jamin_list.append(item)
        elif 'साक्षीदार' in role:
            sakshi_list.append(item)

    # -------------------------------------------------------------
    # प्रकार १: तारण गहाणखत (Mortgage Deed)
    # -------------------------------------------------------------
    if 'तारण' in doc_type or 'गहाण' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("।। श्री ।।\n।। तारण गहाण खत ।।\n").bold = True

        doc.add_paragraph(
            f"तारण गहाण खत {f.get('village_taluka')} येथील मिळकतीचे रक्कम रुपये {f.get('amount_num')} "
            f"({f.get('amount_words')}) चे दि. {f.get('doc_date')}\n"
        )
        
        # बँक (लिहून घेणार)
        doc.add_paragraph(
            f"{f.get('selected_bank')},\n"
            f"तर्फे {f.get('bank_rep_name')}\t\t\t\t\t\t\t\tलिहून घेणार\n"
        )

        # कर्जदार (लिहून देणार)
        p_d = doc.add_paragraph()
        p_d.add_run("कर्जदार (लिहून देणार):\n").bold = True
        for idx, p in enumerate(denar_list):
            p_d.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']} | पॅन: {p['pan']}\n\n")

        # जामीनदार
        if jamin_list:
            p_j = doc.add_paragraph()
            p_j.add_run("जामीनदार:\n").bold = True
            for idx, p in enumerate(jamin_list):
                p_j.add_run(f"{idx+1}) {p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        # अटी व मिळकत वर्णन
        doc.add_paragraph(
            f"कारणे लिहून देतो सतत तारण गहाण खत ऐसाजे कि,\n"
            f"१) मी कर्जदार {f.get('selected_bank')} या संस्थेचा सभासद असून {f.get('bank_meeting_res')} अन्वये "
            f"माझे नांवे रक्कम रुपये {f.get('amount_num')} ({f.get('amount_words')}) चे कर्ज मंजूर झाले आहे.\n"
            f"२) सदर कर्जास व्याजाचा दर द.सा.द.शे. {f.get('bank_interest')} राहील व कर्जाची मुदत {f.get('bank_tenure')} राहील.\n\n"
            f"तारण मिळकतीचे वर्णन :-\n"
            f"{f.get('village_taluka')} येथील गट नं./सि.स.नं. {f.get('gat_no')} एकूण क्षेत्र {f.get('total_area')} पैकी "
            f"तारण दिलेले क्षेत्र {f.get('sold_area')} ची मिळकत सर्व.\n"
            f"चतुःसीमा :-\nपूर्वेस : {f.get('east')} | पश्चिमेस : {f.get('west')} | दक्षिणेस : {f.get('south')} | उत्तरेस : {f.get('north')}\n\n"
            f"येणेप्रमाणे तारण गहाणखत स्वखुषीने लिहून दिले असे सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार २: करारपत्र TO खरेदीपत्र
    # -------------------------------------------------------------
    elif 'करार' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("|| श्री स्वामी समर्थ ||\n|| खरेदीपत्र ||\n").bold = True

        doc.add_paragraph(
            f"रजि. करारपत्र {f.get('karar_ref')} रोजीच्या दस्तावरून खरेदीपत्र {f.get('village_taluka')} येथील मिळकतीचे "
            f"सरकारी बाजारभावाप्रमाणे रक्कम रुपये {f.get('amount_num')} ({f.get('amount_words')}) दि. {f.get('doc_date')} इसवीचे.\n"
        )

        p_g = doc.add_paragraph()
        p_g.add_run("खरेदी घेणार:\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']} | पॅन: {p['pan']}\n\n")

        p_d = doc.add_paragraph()
        p_d.add_run("खरेदी देणार:\n").bold = True
        for p in denar_list:
            p_d.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n")
        if poa_list:
            p_d.add_run(f"यांचे तर्फे वटमुखत्यार म्हणून {poa_list[0]['name']}, {poa_list[0]['age']}\nरा. {poa_list[0]['addr']}\n\n")

        doc.add_paragraph(
            f"मिळकतीचे वर्णन :-\n"
            f"{f.get('village_taluka')} येथील गट नं. {f.get('gat_no')} एकूण क्षेत्र {f.get('total_area')} यापैकी "
            f"क्षेत्र {f.get('sold_area')} ची मिळकत सर्व.\n"
            f"चतुःसीमा :-\nपूर्वेस : {f.get('east')} | पश्चिमेस : {f.get('west')} | दक्षिणेस : {f.get('south')} | उत्तरेस : {f.get('north')}\n\n"
            f"येणेप्रमाणे करारपत्रावरून खरेदीपत्र लिहून दिले असे सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार ३: विनामोबदला हक्कसोडपत्र
    # -------------------------------------------------------------
    elif 'हक्कसोड' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("|| श्री ||\n|| विनामोबदला हक्कसोडपत्र ||\n").bold = True

        doc.add_paragraph(f"{f.get('village_taluka')} येथील एकत्रित मिळकतीचे विनामोबदला हक्कसोडपत्र दि. {f.get('doc_date')} इसवीचे.\n")

        p_g = doc.add_paragraph()
        p_g.add_run("हक्कसोडपत्र लिहून घेणार:\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        p_d = doc.add_paragraph()
        p_d.add_run("हक्कसोडपत्र लिहून देणार:\n").bold = True
        for idx, p in enumerate(denar_list):
            p_d.add_run(f"{idx+1}) {p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        doc.add_paragraph(
            f"मिळकतीचे वर्णन :-\n"
            f"{f.get('village_taluka')} येथील गट नं./सि.स.नं. {f.get('gat_no')} एकूण क्षेत्र {f.get('total_area')} मधील "
            f"हक्कसोड केलेले क्षेत्र {f.get('sold_area')} ची मिळकत सर्व.\n"
            f"चतुःसीमा :-\nपूर्वेस : {f.get('east')} | पश्चिमेस : {f.get('west')} | दक्षिणेस : {f.get('south')} | उत्तरेस : {f.get('north')}\n\n"
            f"आम्ही आमचा वडिलोपार्जित हिस्सा विनामोबदला प्रेमापोटी लिहून घेणार यांचे नावे सोडला आहे. सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार ४: विनामोबदला बक्षीसपत्र
    # -------------------------------------------------------------
    elif 'बक्षीस' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("|| श्री ||\n|| विनामोबदला बक्षीसपत्र ||\n").bold = True

        doc.add_paragraph(f"{f.get('village_taluka')} येथील मिळकतीचे विनामोबदला बक्षीसपत्र रक्कम रुपये {f.get('amount_num')} चे दि. {f.get('doc_date')}.\n")

        p_d = doc.add_paragraph()
        p_d.add_run("बक्षीसपत्र लिहून देणार:\n").bold = True
        for p in denar_list:
            p_d.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        p_g = doc.add_paragraph()
        p_g.add_run("बक्षीसपत्र लिहून घेणार:\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        if manyata_list:
            p_m = doc.add_paragraph()
            p_m.add_run("मान्यता देणार:\n").bold = True
            for p in manyata_list:
                p_m.add_run(f"{p['name']}, {p['age']}, रा. {p['addr']}\n\n")

        doc.add_paragraph(
            f"मिळकतीचे वर्णन :-\n"
            f"{f.get('village_taluka')} येथील मिळकत गट नं./सि.स.नं. {f.get('gat_no')} एकूण क्षेत्र {f.get('total_area')} पैकी "
            f"बक्षीस दिलेले क्षेत्र {f.get('sold_area')} ची मिळकत सर्व.\n"
            f"चतुःसीमा :-\nपूर्वेस : {f.get('east')} | पश्चिमेस : {f.get('west')} | दक्षिणेस : {f.get('south')} | उत्तरेस : {f.get('north')}\n\n"
            f"सदरची मिळकत प्रेमाने स्वखुषीने बक्षीस दिली आहे. सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार ५: चुकदुरुस्ती लेख
    # -------------------------------------------------------------
    elif 'चुकदुरुस्ती' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("|| श्री ||\n|| चुकदुरुस्ती लेख ||\n").bold = True

        doc.add_paragraph(f"चुकदुरुस्ती लेख {f.get('village_taluka')} येथील मिळकतीबाबत तारीख दि. {f.get('doc_date')} इसवीचे.\n")

        p_g = doc.add_paragraph()
        p_g.add_run("लिहून घेणार:\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, रा. {p['addr']}\n")

        p_d = doc.add_paragraph()
        p_d.add_run("\nलिहून देणार:\n").bold = True
        for p in denar_list:
            p_d.add_run(f"{p['name']}, {p['age']}, रा. {p['addr']}\n")

        doc.add_paragraph(
            f"\nमिळकतीचे वर्णन :-\n"
            f"गट नं. {f.get('gat_no')} क्षेत्र {f.get('sold_area')}. चतुःसीमा : पूर्वेस- {f.get('east')}, पश्चिमेस- {f.get('west')}, दक्षिणेस- {f.get('south')}, उत्तरेस- {f.get('north')}.\n\n"
            f"सदर दस्तामध्ये पूर्वी नजरचुकीने झालेली चूक या चुकदुरुस्ती लेखाने दुरुस्त करण्यात आली आहे. बाकी सर्व अटी मूळ दस्ताप्रमाणे कायम राहतील. सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार ६: रद्द न होणारे वटमुखत्यारपत्र
    # -------------------------------------------------------------
    elif 'वटमुखत्यार' in doc_type:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run("|| श्री ||\n|| रद्द न होणारे वटमुखत्यारपत्र ||\n").bold = True

        doc.add_paragraph(f"वटमुखत्यार पत्र {f.get('village_taluka')} येथील मिळकतीचे दिनांक {f.get('doc_date')} रोजीचे.\n")

        p_g = doc.add_paragraph()
        p_g.add_run("लिहून घेणार (मुखत्यार):\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        p_d = doc.add_paragraph()
        p_d.add_run("लिहून देणार (मालक):\n").bold = True
        for p in denar_list:
            p_d.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        doc.add_paragraph(
            f"मिळकतीचे वर्णन :-\n"
            f"गट नं. {f.get('gat_no')} क्षेत्र {f.get('sold_area')}. चतुःसीमा : पूर्वेस- {f.get('east')}, पश्चिमेस- {f.get('west')}, दक्षिणेस- {f.get('south')}, उत्तरेस- {f.get('north')}.\n\n"
            f"सदर मिळकतीची देखभाल, खरेदी-विक्री, निबंधक कार्यालयात हजर राहणे, स्वाक्षरी करणे इत्यादी सर्व अधिकार वटमुखत्यार यांना प्रदान करण्यात आले आहेत. सही दि. {f.get('doc_date')}."
        )

    # -------------------------------------------------------------
    # प्रकार ७: सामान्य / इतर खरेदीखत (प्लॉट, शेती, सीटीएस)
    # -------------------------------------------------------------
    else:
        p_head = doc.add_paragraph()
        p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_head.add_run(f"|| श्री ||\n|| {doc_type} ||\n").bold = True

        doc.add_paragraph(
            f"कायम खुश {doc_type} {f.get('village_taluka')} येथील मिळकतीबाबतचे रक्कम रुपये {f.get('amount_num')} "
            f"({f.get('amount_words')}) दि. {f.get('doc_date')} इसवीचे.\n"
        )

        p_g = doc.add_paragraph()
        p_g.add_run("खरेदी घेणार (लिहून घेणार):\n").bold = True
        for p in ghenar_list:
            p_g.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']} | पॅन: {p['pan']}\n\n")

        p_d = doc.add_paragraph()
        p_d.add_run("खरेदी देणार (लिहून देणार):\n").bold = True
        for p in denar_list:
            p_d.add_run(f"{p['name']}, {p['age']}, धंदा- {p['occ']},\nरा. {p['addr']}\nआ.नं. {p['aadhar']}\n\n")

        if poa_list:
            p_d.add_run(f"यांचे तर्फे वटमुखत्यार म्हणून {poa_list[0]['name']}, {poa_list[0]['age']}\nरा. {poa_list[0]['addr']}\n\n")

        doc.add_paragraph(
            f"मिळकतीचे वर्णन :-\n"
            f"{f.get('village_taluka')} येथील गट नं./सि.स.नं. {f.get('gat_no')} एकूण क्षेत्र {f.get('total_area')} पैकी "
            f"खरेदी क्षेत्र {f.get('sold_area')} ची मिळकत सर्व.\n"
            f"चतुःसीमा :-\nपूर्वेस : {f.get('east')} | पश्चिमेस : {f.get('west')} | दक्षिणेस : {f.get('south')} | उत्तरेस : {f.get('north')}\n\n"
            f"येणेप्रमाणे कायम खुश खरेदीपत्र लिहून दिले असे सही दि. {f.get('doc_date')}."
        )

    # फाईल तयार करून डाऊनलोड देणे
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    clean_doc = doc_type.replace(' ', '_')
    gat_clean = str(f.get('gat_no', '1')).replace('/', '_')
    filename = f"{clean_doc}_Gat_{gat_clean}.docx"

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ७. IGR व संपर्क Routes
@app.route('/igr')
def igr():
    return render_template('igr.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

if __name__ == '__main__':
    app.run(debug=True)

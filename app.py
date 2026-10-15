from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import os
import io
import json

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

# ४. प्लॉट खरेदीपत्र डेटा नोंदणी फॉर्म
@app.route('/plot_draft')
def plot_draft():
    return render_template('plot_form.html')

# ५. करारपत्र TO खरेदीपत्र फॉर्म
@app.route('/karar_to_kharedi_draft')
def karar_to_kharedi_draft():
    return render_template('karar_to_kharedi_form.html')

# ६. सामान्य फॉर्म
@app.route('/draft_form')
def draft_form():
    doc_type = request.args.get('type', '')
    if doc_type == 'करार TO खरेदीपत्र':
        return redirect(url_for('karar_to_kharedi_draft'))
    return render_template('index.html', selected_type=doc_type)

# ७. दस्तऐवज स्कॅनिंग बॅकएंड API
@app.route('/scan_party_document', methods=['POST'])
def scan_party_document():
    # फ्रंटएंडला थेट उत्तर परत देणे
    return jsonify({
        "status": "ready"
    })

# ८. Word (.docx) दस्त निर्मिती (वर्गीकृत पक्षकारांसह)
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
    form = request.form
    doc = Document()

    # मथळा
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("|| श्री ||\n|| खरेदीपत्र ||\n").bold = True

    # प्राथमिक माहिती
    intro = (
        f"तुकडी व जिल्हा {form.get('taluka_dist')} मधील {form.get('village')} या गावाचे नगरपालिका हद्दीतील "
        f"गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये "
        f"{form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {form.get('rate_sqm')} "
        f"रुपये प्रमाणे दि. {form.get('doc_date')} इ.\n\n"
    )
    doc.add_paragraph(intro)

    # पक्षकारांची वर्गवारी (Categorization)
    roles = form.getlist('party_role[]')
    names = form.getlist('party_name[]')
    ages = form.getlist('party_age[]')
    occs = form.getlist('party_occ[]')
    addrs = form.getlist('party_address[]')
    aadhars = form.getlist('party_aadhar[]')
    pans = form.getlist('party_pan[]')
    mobs = form.getlist('party_mobile[]')

    ghenar_list = []
    denar_list = []
    poa_list = []
    manyata_list = []
    sakshi_list = []

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
        elif 'साक्षीदार' in role:
            sakshi_list.append(item)

    # १. लिहून घेणार छापणी
    p_ghenar = doc.add_paragraph()
    p_ghenar.add_run("खरेदी घेणार (लिहून घेणार):\n").bold = True
    for idx, p in enumerate(ghenar_list):
        p_ghenar.add_run(
            f"{idx+1}) {p['name']}, {p['age']}, धंदा- {p['occ']},\n"
            f"{p['addr']}\n"
            f"आ.नं.- {p['aadhar']} | पॅन नं.- {p['pan']} | मो.नं.- {p['mob']}\n\n"
        )

    # २. लिहून देणार छापणी
    p_denar = doc.add_paragraph()
    p_denar.add_run("विरुद्ध\n\nखरेदी देणार (लिहून देणार):\n").bold = True
    for idx, p in enumerate(denar_list):
        p_denar.add_run(
            f"{idx+1}) {p['name']}, {p['age']}, धंदा- {p['occ']},\n"
            f"{p['addr']}\n"
            f"आ.नं.- {p['aadhar']} | पॅन नं.- {p['pan']} | मो.नं.- {p['mob']}\n\n"
        )

    # ३. वटमुखत्यार असल्यास
    if poa_list:
        p_poa = doc.add_paragraph()
        p_poa.add_run("यांचे तर्फे वटमुखत्यार म्हणून:\n").bold = True
        for idx, p in enumerate(poa_list):
            p_poa.add_run(
                f"{idx+1}) {p['name']}, {p['age']}, धंदा- {p['occ']},\n"
                f"{p['addr']}\n"
                f"आ.नं.- {p['aadhar']} | पॅन नं.- {p['pan']} | मो.नं.- {p['mob']}\n\n"
            )

    # ४. मान्यता देणार / संमती देणार असल्यास
    if manyata_list:
        p_m = doc.add_paragraph()
        p_m.add_run("संमती / मान्यता देणार:\n").bold = True
        for idx, p in enumerate(manyata_list):
            p_m.add_run(f"{idx+1}) {p['name']}, {p['age']}, {p['addr']}\n")
        p_m.add_run("\n")

    # मिळकतीचे वर्णन
    milkat = (
        f"कारणे लिहून देतो खरेदीपत्र ऐसाजे की,\n"
        f"1) मिळकतीचे वर्णन:- \n"
        f"जिल्हा व तुकडी कोल्हापूर पोट तुकडी {form.get('taluka_dist')} मे. दुय्यम निबंधकसो यांचे अधिकार क्षेत्रातील "
        f"{form.get('village')} या गावाचे हद्दीतील मिळकत खरेदी देणार यांचे स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे.\n"
        f"री.स.नं./गट नं. {form.get('gat_no')} यामधील खरेदी देणार यांचे मालकीचा गुंठेवारी नियमितीकरण झालेले प्लॉट नंबर "
        f"{form.get('plot_no')} याचे क्षेत्र {form.get('final_area')} ची खुली प्लॉट मिळकत व त्यामधील बांधकाम क्षेत्र "
        f"{form.get('const_area')} ची मिळकत सर्व.\n\n"
        f"यांसी चतुःसिमा पुढीलप्रमाणे :-\n"
        f"पूर्वेस :- {form.get('east')} \n"
        f"पश्चिमेस :- {form.get('west')} \n"
        f"दक्षिणेस :- {form.get('south')} \n"
        f"उत्तरेस :- {form.get('north')} \n\n"
        f"2) मोबदला व भरणा तपशील:-\n"
        f"खरेदी मोबदल्याची एकूण रक्कम रुपये {form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')})\n\n"
        f"3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या हक्कासह सर्व मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\n"
        f"म्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता. {form.get('doc_date')} इ."
    )
    doc.add_paragraph(milkat)

    # साक्षीदार असल्यास
    if sakshi_list:
        p_sakshi = doc.add_paragraph("\nसाक्षीदार :-\n")
        for idx, p in enumerate(sakshi_list):
            p_sakshi.add_run(f"{idx+1}) {p['name']}, रा. {p['addr']}\n")

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    first_ghenar = ghenar_list[0]['name'].replace(' ', '_') if ghenar_list else "Ghenar"
    gat_clean = str(form.get('gat_no', '')).replace('/', '_')
    filename = f"PLOT_Kharedipatra_Gat_{gat_clean}_{first_ghenar}.docx"

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ९. IGR व संपर्क Routes
@app.route('/igr')
def igr():
    return render_template('igr.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

if __name__ == '__main__':
    app.run(debug=True)

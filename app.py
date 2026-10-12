import os
import re
import json
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import pytesseract
from PIL import Image
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def extract_satbara_data(file_path):
    text = ""
    if file_path.lower().endswith('.pdf'):
        try:
            from pdf2image import convert_from_path
            pages = convert_from_path(file_path, first_page=1, last_page=1)
            for page in pages:
                text += pytesseract.image_to_string(page, lang='mar+eng')
        except Exception:
            text = ""
    else:
        img = Image.open(file_path)
        text = pytesseract.image_to_string(img, lang='mar+eng')

    res = {
        'village': '',
        'taluka': '',
        'district': '',
        'gat_no': '',
        'total_area': '0 हे 00 आर 00 पा',
        'potkharaba': '0 हे 00 आर 00 पा',
        'aakar': '0 रुपये 00 पैसे'
    }

    v_match = re.search(r'गाव[:\s]+([^\n,]+)', text)
    if v_match: res['village'] = v_match.group(1).strip()

    t_match = re.search(r'तालुका[:\s]+([^\n,]+)', text)
    if t_match: res['taluka'] = t_match.group(1).strip()

    d_match = re.search(r'जिल्हा[:\s]+([^\n,]+)', text)
    if d_match: res['district'] = d_match.group(1).strip()

    gat_match = re.search(r'(गट\s*क्रमांक|भूमापन\s*क्रमांक|गट\s*नंबर)[:\s]*([0-9\s/अ-हA-Za-z]+)', text)
    if gat_match: res['gat_no'] = gat_match.group(2).strip()

    area_match = re.search(r'([0-9]+)\s*हे\s*([0-9]+)\s*आर', text)
    if area_match:
        res['total_area'] = f"{area_match.group(1)} हे {area_match.group(2)} आर 00 पा"

    pot_match = re.search(r'पोटखराब[^\d]*([0-9]+)\s*हे\s*([0-9]+)\s*आर', text)
    if pot_match:
        res['potkharaba'] = f"{pot_match.group(1)} हे {pot_match.group(2)} आर 00 पा"

    aakar_match = re.search(r'([0-9]+)\s*(?:रुपये|\.)\s*([0-9]+)\s*पैसे', text)
    if aakar_match:
        res['aakar'] = f"{aakar_match.group(1)} रुपये {aakar_match.group(2)} पैसे"

    return res

def extract_aadhaar_front(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang='eng+mar')
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    data = {'name': '', 'dob': '', 'age': '', 'aadhaar_no': '', 'gender': ''}

    aadhaar_match = re.search(r'\b\d{4}\s\d{4}\s\d{4}\b', text)
    if aadhaar_match:
        data['aadhaar_no'] = aadhaar_match.group(0).replace(" ", "")

    dob_match = re.search(r'(DOB|Birth|Year of Birth)[:\s]*([0-9]{2}/[0-9]{2}/[0-9]{4}|[0-9]{4})', text, re.IGNORECASE)
    if dob_match:
        val = dob_match.group(2)
        data['dob'] = val
        if len(val) == 4:
            data['age'] = str(datetime.now().year - int(val))
        elif len(val) == 10:
            birth_year = int(val.split('/')[-1])
            data['age'] = str(datetime.now().year - birth_year)

    if re.search(r'\b(Male|पुरूष)\b', text, re.IGNORECASE):
        data['gender'] = 'पुरूष'
    elif re.search(r'\b(Female|स्त्री)\b', text, re.IGNORECASE):
        data['gender'] = 'स्त्री'

    for line in lines[:5]:
        if not any(k in line.lower() for k in ['government', 'india', 'male', 'female', 'dob', 'आधार', 'भारत', 'father', 'birth']):
            if len(line) > 3 and not re.search(r'\d', line):
                data['name'] = line
                break

    return data

def extract_aadhaar_back(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang='eng+mar')
    data = {'address': ''}

    addr_match = re.search(r'(Address|पत्ता)[:\s]*(.*)', text, re.IGNORECASE | re.DOTALL)
    if addr_match:
        data['address'] = addr_match.group(2).replace('\n', ' ').strip()[:300]
    else:
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        data['address'] = " ".join(lines[:4])

    return data

def extract_pan_data(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img, lang='eng')
    data = {'pan_no': ''}

    pan_match = re.search(r'[A-Z]{5}[0-9]{4}[A-Z]{1}', text)
    if pan_match:
        data['pan_no'] = pan_match.group(0)

    return data

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/presentation')
def presentation():
    district = request.args.get('drpDistrictNew', '')
    taluka = request.args.get('drpTalukaNew', '')
    village = request.args.get('txtVillageNew', '')
    return render_template('presentation.html', district=district, taluka=taluka, village=village)

@app.route('/property_details')
def property_details():
    district = request.args.get('district', '')
    taluka = request.args.get('taluka', '')
    village = request.args.get('village', '')
    token_no = request.args.get('token_no', '')
    article = request.args.get('cboArticle', '')
    article_desc = request.args.get('cboArticleDescription', '')
    market_val = request.args.get('txtMarketValue', '')
    consideration_amt = request.args.get('txtConsiderationAmount', '')
    presentation_date = request.args.get('txtDateOfPresentation', '')
    return render_template('property.html', district=district, taluka=taluka, village=village, 
                           token_no=token_no, article=article, article_desc=article_desc,
                           market_val=market_val, consideration_amt=consideration_amt,
                           presentation_date=presentation_date)

@app.route('/scan_satbara', methods=['POST'])
def scan_satbara_route():
    if 'file' not in request.files:
        return jsonify({'error': 'फाईल नाही'}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
    f.save(path)
    try:
        res = extract_satbara_data(path)
    finally:
        if os.path.exists(path):
            os.remove(path)
    return jsonify(res)

@app.route('/party_entry')
def party_entry():
    district = request.args.get('district', '')
    taluka = request.args.get('taluka', '')
    village = request.args.get('village', '')
    token_no = request.args.get('token_no', '')
    article = request.args.get('article', '')
    article_desc = request.args.get('article_desc', 'खरेदीपत्र')
    market_val = request.args.get('market_val', '')
    consideration_amt = request.args.get('consideration_amt', '')
    presentation_date = request.args.get('presentation_date', '')
    properties_data = request.args.get('properties_data', '[]')
    return render_template('index.html', district=district, taluka=taluka, village=village, 
                           token_no=token_no, article=article, article_desc=article_desc,
                           market_val=market_val, consideration_amt=consideration_amt,
                           presentation_date=presentation_date, properties_data=properties_data)

@app.route('/scan_aadhaar_front', methods=['POST'])
def scan_aadhaar_front_route():
    if 'file' not in request.files:
        return jsonify({'error': 'फाईल नाही'}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
    f.save(path)
    res = extract_aadhaar_front(path)
    os.remove(path)
    return jsonify(res)

@app.route('/scan_aadhaar_back', methods=['POST'])
def scan_aadhaar_back_route():
    if 'file' not in request.files:
        return jsonify({'error': 'फाईल नाही'}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
    f.save(path)
    res = extract_aadhaar_back(path)
    os.remove(path)
    return jsonify(res)

@app.route('/scan_pan', methods=['POST'])
def scan_pan_route():
    if 'file' not in request.files:
        return jsonify({'error': 'फाईल नाही'}), 400
    f = request.files['file']
    path = os.path.join(app.config['UPLOAD_FOLDER'], f.filename)
    f.save(path)
    res = extract_pan_data(path)
    os.remove(path)
    return jsonify(res)

@app.route('/generate_docx', methods=['POST'])
def generate_docx():
    parties = json.loads(request.form.get('parties_data', '[]'))
    properties = json.loads(request.form.get('properties_data', '[]'))
    payments = json.loads(request.form.get('payments_data', '[]'))
    district = request.form.get('district', '')
    taluka = request.form.get('taluka', '')
    village = request.form.get('village', '')
    article_desc = request.form.get('article_desc', '').strip() or 'खरेदीपत्र'
    consideration_amt = request.form.get('consideration_amt', '0')
    presentation_date = request.form.get('presentation_date', '')

    purchasers = [p for p in parties if 'घेणार' in p.get('party_type', '')]
    sellers = [p for p in parties if 'देणार' in p.get('party_type', '')]
    if not purchasers and parties: purchasers = [parties[0]]
    if not sellers and len(parties) > 1: sellers = parties[1:]

    first_prop = properties[0] if properties else {}
    gat_no = first_prop.get('attrNo', '')
    prop_desc_text = first_prop.get('desc', '')

    purchaser_name = purchasers[0].get('party_name', 'पक्षकार') if purchasers else 'पक्षकार'

    clean_gat = re.sub(r'[\\/*?:"<>|]', '_', str(gat_no))
    clean_purchaser = re.sub(r'[\\/*?:"<>|]', '_', str(purchaser_name))
    file_name = f"{article_desc} मौजे {village} गट नंबर {clean_gat} आणि {clean_purchaser}.docx"
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)

    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    # १. मथळा
    p_top = doc.add_paragraph()
    p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_shree = p_top.add_run('|| श्री ||\n')
    r_shree.font.size = Pt(14)
    r_shree.font.bold = True
    
    r_title = p_top.add_run(f'|| {article_desc} ||\n')
    r_title.font.size = Pt(18)
    r_title.font.bold = True

    # २. प्रास्ताविक
    p_head = doc.add_paragraph()
    p_head.paragraph_format.line_spacing = 1.15
    p_head.add_run(
        f"तुकडी व जिल्हा {district} पोट तुकडी तहसील {taluka} मधील मौजे {village} या गावाचे हद्दीतील मिळकतीचे "
        f"रक्कम रुपये {consideration_amt}/- दि. {presentation_date} इ.\n"
    )

    # ३. लिहून घेणार तपशील
    for pur in purchasers:
        p_pur = doc.add_paragraph()
        p_pur.paragraph_format.line_spacing = 1.15
        p_pur.add_run(f"{pur.get('party_name', '')}, वय {pur.get('party_age', '')} वर्ष, धंदा- घरकाम/व्यवसाय, ")
        p_pur.add_run(f"रा. {pur.get('party_address', '')} ")
        r_tag1 = p_pur.add_run("............... लिहून घेणार\n")
        r_tag1.bold = True
        pan_val = pur.get('pan_no', '') or 'उपलब्ध नाही'
        mob_val = pur.get('mobile_no', '') or 'उपलब्ध नाही'
        p_pur.add_run(f"ओळख क्रमांक: [आधार क्रमांक: ________________]  पॅन नं.: {pan_val}  मोबाईल नं.: {mob_val}")

    doc.add_paragraph()

    # ४. लिहून देणार तपशील
    for sel in sellers:
        p_sel = doc.add_paragraph()
        p_sel.paragraph_format.line_spacing = 1.15
        p_sel.add_run(f"{sel.get('party_name', '')}, वय {sel.get('party_age', '')} वर्ष, धंदा- व्यवसाय, ")
        p_sel.add_run(f"रा. {sel.get('party_address', '')} ")
        r_tag2 = p_sel.add_run("............... लिहून देणार\n")
        r_tag2.bold = True
        pan_val = sel.get('pan_no', '') or 'उपलब्ध नाही'
        mob_val = sel.get('mobile_no', '') or 'उपलब्ध नाही'
        p_sel.add_run(f"ओळख क्रमांक: [आधार क्रमांक: ________________]  पॅन नं.: {pan_val}  मोबाईल नं.: {mob_val}")

    doc.add_paragraph('--------------------------------------------------------------------------------')

    # ५. कारणे मजकूर
    p_reason = doc.add_paragraph()
    p_reason.add_run(f"कारणे लिहून देतो {article_desc} ऐसाजे की,\n").bold = True

    # कलम १: मिळकतीचे वर्णन
    p_c1 = doc.add_paragraph()
    p_c1.add_run("१) मिळकतीचे वर्णन :-\n").bold = True
    p_c1.add_run(
        f"जिल्हा व तुकडी {district} पोट तुकडी ता. {taluka} मे. दुय्यम निबंधक सो. {taluka} यांचे अधिकार क्षेत्रातील "
        f"मौजे {village} या गावाचे हद्दीतील मिळकत देणार यांचे स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे. "
        f"त्याचे वर्णन खालीलप्रमाणे :-\n"
    )
    p_c1.add_run(f"{prop_desc_text}\n\n")
    p_c1.add_run(
        "येणेप्रमाणे वर नमूद केलेली मिळकत यातील देणार यांचे स्वसंपादित मालकी वहिवाटीची मिळकत आहे. "
        "उपरोक्त मिळकतीस सोयीसाठी संक्षिप्त रुपात “सदर मिळकत” असे संबोधले आहे."
    )

    # कलम २: मोबदला व ताबा
    p_c2 = doc.add_paragraph()
    p_c2.add_run("२) मोबदला व ताबा :-\n").bold = True
    p_c2.add_run(
        f"सदर मिळकतीचा व्यवहार उभयतांच्या विचारांती ठरवून त्याप्रमाणे रक्कम रुपये {consideration_amt}/- "
        f"घेणार यांचेकडून देणार यांनी पूर्णपणे स्वीकारून वरील वर्णन केलेल्या मिळकतीचा प्रत्यक्ष ताबा घेणार यांचे स्वाधीन केला आहे.\n"
    )

    # भरणा तपशील तक्ता (Dynamic Multiple Payments)
    p_tbl_lbl = doc.add_paragraph()
    p_tbl_lbl.add_run("भरणा तपशील :-").bold = True
    tbl_pay = doc.add_table(rows=1, cols=4)
    tbl_pay.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_pay.style = 'Table Grid'
    hdr = tbl_pay.rows[0].cells
    hdr[0].text = "अ.नं."
    hdr[1].text = "भरणा प्रकार"
    hdr[2].text = "रक्कम (₹)"
    hdr[3].text = "क्रमांक, दिनांक व बँक तपशील"

    if payments:
        for p_idx, py in enumerate(payments, start=1):
            pr = tbl_pay.add_row().cells
            pr[0].text = str(p_idx)
            pr[1].text = str(py.get('mode', ''))
            pr[2].text = f"{py.get('amount', 0):,}/-"
            pr[3].text = f"क्र./Ref: {py.get('refNo', '-')} दि. {py.get('date', '-')} तपशील: {py.get('details', '-')}"
        
        # एकूण रो
        tot_r = tbl_pay.add_row().cells
        tot_r[0].text = "एकूण"
        tot_r[1].text = "एकूण रक्कम"
        tot_r[2].text = f"{consideration_amt}/-"
        tot_r[3].text = "येणेप्रमाणे खरेदी मोबदल्याची रक्कम पूर्ण पोहचली असून भरण्याबाबत कसलीही तक्रार नाही."
    else:
        pr = tbl_pay.add_row().cells
        pr[0].text = "1"
        pr[1].text = "एकरकमी"
        pr[2].text = f"{consideration_amt}/-"
        pr[3].text = "सदर मोबदला रक्कम रोख / चेकद्वारे पूर्ण पोहचली आहे."

    doc.add_paragraph()

    # इतर सर्व कायदेशीर कलमे
    clauses = [
        ("३)", f"या {article_desc}ने आजपासून घेणार हे वर वर्णन केलेल्या मिळकतीचे कायदेशीर मालक व कब्जेदार झाले आहेत. सदर मिळकत सर्व प्रकारच्या बोजांपासून मुक्त व निर्वेध आहे."),
        ("४)", "या मिळकतीसंबंधाने येथून पुढे भरावयाचा सर्व सरकारी कर, पट्टी, फाळा घेणार यांनी भरावयाचा आहे. मागील सर्व देणी देणार यांनी भागविलेली आहेत."),
        ("५)", f"या {article_desc}च्या आधारे मिळकतीचे अभिलेख दफ्तरी घेणार यांनी आपले नावे करून घ्यावे. त्यावेळी जरूर ती संमती देणार विनामोबदला देतील."),
        ("६)", "सदर मिळकतीबाबत अन्य कोणाशीही करारमदार अथवा व्यवहार केलेला नाही. असा काही व्यवहार केला असल्यास तो या दस्ताने रद्द समजावा."),
        ("७)", "सदर मिळकत शासन अथवा स्थानिक प्राधिकरणाने कोणत्याही प्रकल्पासाठी आरक्षित केलेली नाही, तसेच भूसंपादनाची कोणतीही नोटीस आलेली नाही."),
        ("८)", "सदर मिळकतीवर कोणत्याही न्यायालयात वाद चालू नाही अथवा जप्ती आदेश नाही. भविष्यात कोणताही वाद उद्भवल्यास त्याचे निवारण देणार स्वखर्चाने करून देतील."),
        ("९)", f"घेणार यांनी मिळकतीची व कागदपत्रांची संपूर्ण खात्री करून घेतलेली आहे. हे {article_desc} उभयतांनी स्वसंतोषाने व अक्कलहुशारीने लिहून दिले व घेतले असे.")
    ]

    for num, txt in clauses:
        p_cl = doc.add_paragraph()
        p_cl.paragraph_format.line_spacing = 1.15
        p_cl.add_run(f"{num} ").bold = True
        p_cl.add_run(txt)

    p_end = doc.add_paragraph()
    p_end.add_run(f"\nम्हणून स्वसंतोषाने लिहून दिले कायम {article_desc} असे सही. दिनांक: {presentation_date} इ.\n").bold = True

    # स्वाक्षरी, फोटो व अंगठ्याचा तक्ता
    tbl_sign = doc.add_table(rows=1, cols=3)
    tbl_sign.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sign.style = 'Table Grid'
    shdr = tbl_sign.rows[0].cells
    shdr[0].text = "पक्षकाराचे नाव व पत्ता"
    shdr[1].text = "फोटो"
    shdr[2].text = "अंगठ्याचा ठसा व सही"

    for pur in purchasers:
        sr = tbl_sign.add_row().cells
        sr[0].text = f"लिहून घेणार:\n{pur.get('party_name', '')}\nरा. {pur.get('party_address', '')}\n\nसही: ....................."
        sr[1].text = "\n\n[ फोटो ]\n\n"
        sr[2].text = "\n\n[ डाव्या हाताचा अंगठा ]\n\n"

    for sel in sellers:
        sr2 = tbl_sign.add_row().cells
        sr2[0].text = f"लिहून देणार:\n{sel.get('party_name', '')}\nरा. {sel.get('party_address', '')}\n\nसही: ....................."
        sr2[1].text = "\n\n[ फोटो ]\n\n"
        sr2[2].text = "\n\n[ डाव्या हाताचा अंगठा ]\n\n"

    doc.save(file_path)
    return send_file(file_path, as_attachment=True, download_name=file_name)

if __name__ == '__main__':
    print("Starting NK CONSULTANCY Server...")
    app.run(host='0.0.0.0', port=5000, debug=True)
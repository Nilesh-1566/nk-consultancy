from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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

# ४. प्लॉट खरेदीपत्र फॉर्म
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

# ७. प्लॉट खरेदीपत्र Word निर्मिती
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
    form = request.form
    doc = Document()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.add_run("|| श्री ||\n|| खरेदीपत्र ||\n").bold = True

    intro = (
        f"तुकडी व जिल्हा {form.get('taluka_dist')} मधील {form.get('village')} या गावाचे नगरपालिका हद्दीतील "
        f"गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये "
        f"{form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {form.get('rate_sqm')} "
        f"रुपये प्रमाणे दि. {form.get('doc_date')} इ.\n\n"
    )
    doc.add_paragraph(intro)

    g_names = form.getlist('ghenar_name[]')
    g_ages = form.getlist('ghenar_age_occ[]')
    g_addresses = form.getlist('ghenar_address[]')
    g_aadhars = form.getlist('ghenar_aadhar[]')
    g_pans = form.getlist('ghenar_pan[]')
    g_mobs = form.getlist('ghenar_mobile[]')

    p_ghenar = doc.add_paragraph()
    p_ghenar.add_run("लिहून घेणार (खरेदी घेणारा):\n").bold = True
    for i in range(len(g_names)):
        p_ghenar.add_run(
            f"{i+1}) {g_names[i]}, {g_ages[i] if i < len(g_ages) else ''}\n"
            f"{g_addresses[i] if i < len(g_addresses) else ''}\n"
            f"आ.क्र. {g_aadhars[i] if i < len(g_aadhars) else ''} | पॅन: {g_pans[i] if i < len(g_pans) else ''} | मो.: {g_mobs[i] if i < len(g_mobs) else ''}\n\n"
        )

    d_names = form.getlist('denar_name[]')
    d_ages = form.getlist('denar_age_occ[]')
    d_addresses = form.getlist('denar_address[]')
    d_aadhars = form.getlist('denar_aadhar[]')
    d_pans = form.getlist('denar_pan[]')
    d_mobs = form.getlist('denar_mobile[]')

    p_denar = doc.add_paragraph()
    p_denar.add_run("विरुद्ध\n\nलिहून देणार (खरेदी देणारा):\n").bold = True
    for i in range(len(d_names)):
        p_denar.add_run(
            f"{i+1}) {d_names[i]}, {d_ages[i] if i < len(d_ages) else ''}\n"
            f"{d_addresses[i] if i < len(d_addresses) else ''}\n"
            f"आ.क्र. {d_aadhars[i] if i < len(d_aadhars) else ''} | पॅन: {d_pans[i] if i < len(d_pans) else ''} | मो.: {d_mobs[i] if i < len(d_mobs) else ''}\n\n"
        )

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
        f"3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\n"
        f"म्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता. {form.get('doc_date')} इ."
    )
    doc.add_paragraph(milkat)

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    gat_no_clean = str(form.get('gat_no', '87')).replace('/', '_').replace(' ', '_')
    plot_no_clean = str(form.get('plot_no', '26')).replace('/', '_').replace(' ', '_')
    first_ghenar = str(g_names[0]).strip().replace(' ', '_') if g_names else "Ghenar"
    filename = f"PLOT_Kharedipatra_Gat_{gat_no_clean}_Plot_{plot_no_clean}_{first_ghenar}.docx"

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# ८. करारपत्र TO खरेदीपत्र Word निर्मिती (नवीन अचूक फॉरमॅट)
@app.route('/generate_karar_to_kharedi_docx', methods=['POST'])
def generate_karar_to_kharedi_docx():
    f = request.form
    doc = Document()

    # मथळा
    p_head = doc.add_paragraph()
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p_head.add_run("|| श्री स्वामी समर्थ ||\n|| खरेदीपत्र ||\n")
    r1.bold = True
    r1.font.size = Pt(14)

    # प्राथमिक माहिती
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.15
    p_intro.add_run(
        f"रजि. करारपत्र दस्त नं.{f.get('karar_dast_no')}, दि.{f.get('karar_date')} इ.रोजीच्या दस्तावरून खरेदीपत्र "
        f"{f.get('village_taluka')} येथील मिळकतीचे सरकारी बाजारभावाप्रमाणे रक्कम रुपये {f.get('amount_num')} "
        f"(अक्षरी रक्कम रुपये {f.get('amount_words')}) दि. {f.get('doc_date')} इसवीचे.\n\n"
    )

    # खरेदी घेणार
    p_ghenar = doc.add_paragraph()
    p_ghenar.add_run(
        f"{f.get('ghenar_name')}, {f.get('ghenar_age_occ')},\n"
        f"{f.get('ghenar_address')}\t\t\t\t\t\t\t\tखरेदी घेणार\n"
        f"आ.नं.- {f.get('ghenar_aadhar')}  पॅन नं.- {f.get('ghenar_pan')}\n\n"
    )

    # खरेदी देणार व वटमुखत्यार
    p_denar = doc.add_paragraph()
    p_denar.add_run(
        f"{f.get('denar_name')}, {f.get('denar_age_occ')},\n"
        f"{f.get('denar_address')}\n"
        f"आ.नं.- {f.get('denar_aadhar')}\t\t\t\t\t\t\t\tखरेदी देणार\n"
        f"        यांचे तर्फे वटमुखत्यार म्हणून \n"
        f"{f.get('poa_name')}, {f.get('poa_age_occ')},\n"
        f"{f.get('poa_address')}\n"
        f"आ.नं.- {f.get('poa_aadhar')}  पॅन नं.- {f.get('poa_pan')}\n\n"
        f"कारणे मी लिहून देतो करारपत्रावरून खरेदीपत्र ऐसाजे की,\n"
    )

    # भरणा तपशील तक्ता
    doc.add_paragraph("भरणा तपशील :").runs[0].bold = True
    pay_amts = f.getlist('pay_amt[]')
    pay_words = f.getlist('pay_words[]')
    pay_banks = f.getlist('pay_bank[]')
    pay_modes = f.getlist('pay_mode[]')

    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    hdr = table.rows[0].cells
    hdr[0].text = "अ.नं."
    hdr[1].text = "रक्कम रुपये"
    hdr[2].text = "तपशील"

    total_calc = 0
    for i in range(len(pay_amts)):
        row = table.add_row().cells
        row[0].text = str(i + 1)
        row[1].text = pay_amts[i]
        row[2].text = (
            f"अक्षरी रुपये {pay_words[i] if i < len(pay_words) else ''}. खरेदी घेणार यांनी खरेदी देणार यांचे नांवे असणाऱ्या "
            f"{pay_banks[i] if i < len(pay_banks) else ''} बँकमध्ये {pay_modes[i] if i < len(pay_modes) else ''} करारपत्र दस्त नं.{f.get('karar_dast_no')}, "
            f"दि.{f.get('karar_date')} इ.रोजी करारपत्रामध्ये नमूद केले भरणेप्रमाणे घेतलेला आहे. ते आम्हाला मान्य व कबूल आहे."
        )

    # एकूण भरणा
    row_tot = table.add_row().cells
    row_tot[0].text = "एकूण"
    row_tot[1].text = f.get('amount_num')
    row_tot[2].text = f"अक्षरी रुपये {f.get('amount_words')}."

    # मिळकत वर्णन व अटी
    p_body = doc.add_paragraph()
    p_body.paragraph_format.line_spacing = 1.15
    p_body.add_run(
        f"\nयेणे प्रमाणे वर दिलेल्या भरण्याप्रमाणे वर नमूद रुपयाच्या मोबदल्यात तुम्हांस {f.get('village_taluka')} येथील मिळकत करारपत्रावरून खरेदीपत्र लिहून देतो ती {f.get('dist_division')}, "
        f"पोट तुकडी तालुका {f.get('sub_registrar')} मे. दुय्यम निबंधकसो, {f.get('sub_registrar')} यांचेकडील दस्त नं.{f.get('karar_dast_no')}, दि.{f.get('karar_date')} इ.रोजीचे करारपत्राने लिहून दिलेली मिळकत आजरोजी या दस्ताने लिहून देतो. तिचे वर्णन खालीलप्रमाणे :-\n"
        f"\tगट नंबर {f.get('gat_no')} क्षेत्र {f.get('total_area')} या पुरे नंबरची जमीन असून यापैकी क्षेत्र {f.get('sold_area')} ची मिळकत सर्व. यासी चतु:सीमा :-\n"
        f"पूर्वेस :- {f.get('east')} \n"
        f"पश्चीमेस :- {f.get('west')} \n"
        f"दक्षिणेस :- {f.get('south')} \n"
        f"उत्तरेस :- {f.get('north')} \n"
        f"\tयेणेप्रमाणे चतु:सीमेतील मिळकत तदंगभूत वस्तूसह व इजमेंटरीच्या हक्कासह सर्व.\n\n"
        f"\tयेणेप्रमाणे वरील चतु:सीमेतील मिळकत मी तुम्हांस मे.दुय्यम निबंधकसो, {f.get('sub_registrar')} यांचेकडील दस्त नं. {f.get('karar_dast_no')}, दि.{f.get('karar_date')} इ. रोजीच्या करारपत्राने रक्कम रुपये {f.get('amount_num')} रुपयास कायम व खूष खरेदी देणेचे रजि. करारपत्र करून दिला असून सदरची मिळकतीचा कब्जा आजरोजी शांततामय असले स्थितीत कबजेत वहिवाटीस दिलेला आहे. तरी सदर तुम्हांस खरेदी दिलेल्या मिळकतीचा सरकारी आकार फाळा, फंड, शिक्षण कर, जो बसेल तो तुमचा तुम्ही परभारे मालकीने देत जाऊन सदर मिळकतीचा वहिवाट तुम्ही व तुमचे पुत्र पौत्रादी, वंशपरंपरा, निरंतर, निधी-निक्षेप, जलतरू, काष्ठ-पाषाण, तदंगभूत वस्तूसह व इजमेंटचे हक्कासह कायमपणे मालकीने वहिवाट करावी. आता सदर तुम्हांस खरेदी दिले मिळकतीवर माझा अगर माझे हरएक प्रकारचे वारसांचा हक्क, दावा अगर संबंध बिलकुल राहिलेला नाही. सर्व प्रकारे संपूर्ण मालकी तुमची झालेली आहे. मिळकतीसंबंधी सरकारी, सावकारी, भाऊबंदकी वगैरे कोणतेही प्रकारचे लागलिगाड आलेस किंवा खाल खटला आलेस ते मी माझे स्वखर्चाने परस्पर निवारण करून देऊ. तुम्हांस त्याची तोशिष लागू देणार नाही. मिळकतीचे खाते या लेखा आधारे तुम्ही तुमच्या नांवे करून घ्यावे. त्याबाबत आवश्यक त्या ठिकाणी लागणाऱ्या सह्या मी माझे खर्चाने येऊन देणार आहे. त्याची तोशिष लागू देणार नाही. सदरची मिळकत निर्वेध व निष्कर्जी आहे. या लेखास अनुसरून काही दुरुस्ती लेख, प्रतिज्ञापत्र वगैरे करून द्यावे लागलेस ते मी विना मोबदला करून देईन. मी तुम्हांस तुमच्या नांवे करारपत्राप्रमाणे कायमखूष खरेदी करून दिलेली आहे. भरणा सदरी लिहिले प्रमाणे मला मान्य व कबूल असून माझी कोणत्याही प्रकारची तक्रार किंवा हरकत नाही. सदरची मिळकत आजरोजी निर्वेध व निष्कर्जी असून पूर्ण रयतावा आहे.\n\n"
        f"\tसदर मिळकतीचे सिद्धगणाप्रमाणे रक्कम रुपये {f.get('amount_num')} इतकी झाली असून रजि.करार दस्त नं. {f.get('karar_dast_no')}, दि.{f.get('karar_date')} इ.रोजी रक्कम रुपये {f.get('amount_num')} मु.शु.र.रु.{f.get('old_stamp')} व नोंदणी फी १% रक्कम रुपये {f.get('old_reg_fee')} भरलेली आहे. आजरोजी मु.शु.{f.get('new_stamp')} व नोंदणी फी {f.get('new_reg_fee')} चा लावून सदर दस्त पुरा केलेला आहे.\n\n"
        f"\tसदरचे खरेदीपत्र मी सांगितलेप्रमाणे तयार केले असून ते मी अव्वलपासून अखेरपर्यंत वाचून पहिले असून सदर दस्तामधील मजकूर मला मान्य व कबूल आहे. म्हणून मी साक्षीदारांच्या समक्ष माझी स्वाक्षरी करून सदरचे खरेदीपत्र पूर्ण केले आहे. सदरचे खरेदीपत्र मला व माझे सर्व प्रकारच्या इस्टेट वारसांस लागू असून बंधनकारक आहे.\n\n"
        f"\tम्हणून, लिहून दिले स्वखुषीने, अक्कल हुशारीने खरेदीपत्र सही तारीख :- {f.get('doc_date')} इ.\n"
    )

    # स्वाक्षरी तक्ता
    doc.add_paragraph("खरेदीपत्र लिहून घेणार :").runs[0].bold = True
    t_g = doc.add_table(rows=2, cols=4)
    t_g.style = 'Table Grid'
    h_g = t_g.rows[0].cells
    h_g[0].text = "अ.नं."
    h_g[1].text = "नांव"
    h_g[2].text = "छायाचित्र"
    h_g[3].text = "सही / अंगठा"
    r_g = t_g.rows[1].cells
    r_g[0].text = "०१"
    r_g[1].text = f"{f.get('ghenar_name')}"

    doc.add_paragraph("\nखरेदीपत्र लिहून देणार :").runs[0].bold = True
    t_d = doc.add_table(rows=2, cols=4)
    t_d.style = 'Table Grid'
    h_d = t_d.rows[0].cells
    h_d[0].text = "अ.नं."
    h_d[1].text = "नांव"
    h_d[2].text = "छायाचित्र"
    h_d[3].text = "सही / अंगठा"
    r_d = t_d.rows[1].cells
    r_d[0].text = "०१"
    r_d[1].text = f"{f.get('denar_name')} यांचे तर्फे वटमुखत्यार म्हणून {f.get('poa_name')}"

    p_sakshi = doc.add_paragraph("\n\nसाक्षीदार :-\n1) ……………………………………………\t\t2) ……………………………………………")

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    ghenar_clean = str(f.get('ghenar_name', 'Ghenar')).replace(' ', '_').replace('.', '')
    gat_clean = str(f.get('gat_no', '')).replace('/', '_')
    filename = f"KARARPATRA_TO_KHAREDIPATRA_Gat_{gat_clean}_{ghenar_clean}.docx"

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

from flask import Flask, render_template, request, send_file, jsonify, redirect, url_for
from docx import Document
from google import genai
from PIL import Image
import os
import io
import json
import urllib.parse

app = Flask(__name__)

# १. Gemini API Key कॉन्फिगरेशन
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AQ.Ab8RN6LqhXJUZBYHwOTyxDAeTHuGRfCbhk-YiPdeOlj62xdc7g")
client = genai.Client(api_key=GEMINI_API_KEY)


# २. AI द्वारे आधार व पॅन कार्ड स्कॅन करण्याचे Route
@app.route('/scan_card_ai', methods=['POST'])
def scan_card_ai():
    try:
        if 'card_files' not in request.files:
            return jsonify({'error': 'कोणताही फोटो निवडलेला नाही'}), 400

        files = request.files.getlist('card_files')
        if not files or files[0].filename == '':
            return jsonify({'error': 'फोटो फाईल रिकामी आहे'}), 400

        images_payload = []
        for f in files:
            img = Image.open(f.stream)
            images_payload.append(img)

        prompt = """
        या आधार कार्ड किंवा पॅन कार्डच्या फोटोंचे काळजीपूर्वक वाचन करा.
        खालील माहिती अचूकपणे ओळखून फक्त वैध JSON फॉरमॅटमध्ये द्या:
        {
          "name": "व्यक्तीचे संपूर्ण नाव शुद्ध देवनागरी मराठीत (उदा. आनंदी पांडुरंग नामदे)",
          "age": "व्यक्तीचे वय (उदा. वय 63 वर्ष, धंदा- घरकाम)",
          "address": "आधार कार्डच्या मागील बाजूचा संपूर्ण पत्ता शुद्ध मराठीत पिनकोडसह (उदा. W/O: पांडुरंग भाऊ नामदे, नरंदे मार्ग, वठार तर्फ उदगांव, वडगाव, कोल्हापूर, महाराष्ट्र, पिन कोड- 416112)",
          "aadhar": "12 अंकी आधार क्रमांक स्पेससह (उदा. xxxx xxxx xxxx)",
          "pan": "10 अंकी इंग्रजी पॅन क्रमांक (पॅन कार्ड असल्यास)",
          "mobile": "मोबाईल नंबर (कार्डवर आढळल्यास)"
        }
        महत्त्वाचे: फक्त आणि फक्त JSON आउटपुट द्या. इतर कोणताही मजकूर किंवा स्पष्टीकरण लिहू नका.
        """

        # सक्रिय Gemini 2.5 Flash मॉडेल
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[prompt] + images_payload
        )

        clean_text = response.text.replace('```json', '').replace('```', '').strip()
        data = json.loads(clean_text)

        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ३. मुख्य होमपेज
@app.route('/')
def home():
    return render_template('home.html')


# ४. दस्त मसुदा मेनू
@app.route('/draft')
def draft():
    return render_template('draft_menu.html')


# ५. खरेदीखत उप-मेनू
@app.route('/kharedi_menu')
def kharedi_menu():
    return render_template('kharedi_menu.html')


# ६. प्लॉट खरेदीपत्र डेटा नोंदणी फॉर्म
@app.route('/plot_draft')
def plot_draft():
    return render_template('plot_form.html')


# ७. सामान्य फॉर्म
@app.route('/draft_form')
def draft_form():
    doc_type = request.args.get('type', '')
    return render_template('index.html', selected_type=doc_type)


# ८. Word (.docx) दस्त निर्मिती व डाऊनलोड
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
    form = request.form
    doc = Document()

    # शीर्षक
    p_title = doc.add_paragraph()
    p_title.alignment = 1
    p_title.add_run("|| श्री ||\n|| खरेदीपत्र ||\n").bold = True

    # प्राथमिक परिच्छेद
    intro = (
        f"तुकडी व जिल्हा {form.get('taluka_dist')} मधील {form.get('village')} या गावाचे नगरपालिका हद्दीतील "
        f"गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये "
        f"{form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {form.get('rate_sqm')} "
        f"रुपये प्रमाणे दि. {form.get('doc_date')} इ.\n\n"
    )
    doc.add_paragraph(intro)

    # १. लिहून घेणार तपशील
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

    # २. लिहून देणार तपशील
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

    # ३. संमती / मान्यता देणार (असल्यास)
    s_names = form.getlist('samati_name[]')
    s_ages = form.getlist('samati_age_occ[]')
    s_addresses = form.getlist('samati_address[]')
    s_aadhars = form.getlist('samati_aadhar[]')
    s_pans = form.getlist('samati_pan[]')
    s_mobs = form.getlist('samati_mobile[]')

    if s_names and any(s.strip() for s in s_names):
        p_samati = doc.add_paragraph()
        p_samati.add_run("संमती / मान्यता देणार:\n").bold = True
        for i in range(len(s_names)):
            if s_names[i].strip():
                p_samati.add_run(
                    f"{i+1}) {s_names[i]}, {s_ages[i] if i < len(s_ages) else ''}\n"
                    f"{s_addresses[i] if i < len(s_addresses) else ''}\n"
                    f"आ.क्र. {s_aadhars[i] if i < len(s_aadhars) else ''} | पॅन: {s_pans[i] if i < len(s_pans) else ''} | मो.: {s_mobs[i] if i < len(s_mobs) else ''}\n\n"
                )

    # ४. मिळकतीचे वर्णन
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
        f"येणेप्रमाणे चतुःसिमेतील मिळकत सर्व व या मिळकतीमध्ये असलेली जल, तरू, काष्ट, पाषाण, निधी-निक्षेप, तसेच जाणे-येणेच्या हक्कासह, आदी तदंगभूत "
        f"वस्तूसहची गुंठेवारी नियमितीकरण झालेली खुली मिळकत व त्यातील रो बंगलो मिळकत हि या खरेदीपत्राचा विषय आहे.\n\n"
        f"2) मोबदला व भरणा तपशील:-\n"
        f"खरेदी मोबदल्याची एकूण रक्कम रुपये {form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) असून त्याचा तपशील खालीलप्रमाणे:\n"
    )
    doc.add_paragraph(milkat)

    # ५. भरणा नोंदी
    pay_amts = form.getlist('pay_amt[]')
    pay_modes = form.getlist('pay_mode[]')
    pay_banks = form.getlist('pay_bank[]')
    p_pay = doc.add_paragraph()
    for i in range(len(pay_amts)):
        if pay_amts[i].strip():
            p_pay.add_run(f"({i+1}) रक्कम रुपये {pay_amts[i]} - {pay_modes[i] if i < len(pay_modes) else ''} द्वारे {pay_banks[i] if i < len(pay_banks) else ''}. भरणा पावती निराळी नाही.\n")

    doc.add_paragraph(
        f"\n3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या हक्कासह सर्व मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\n"
        f"म्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता. {form.get('doc_date')} इ."
    )

    # Word फाईल तयार करून डाऊनलोड करणे
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


# ९. IGR व संपर्क Routes
@app.route('/igr')
def igr():
    return render_template('igr.html')


@app.route('/contact')
def contact():
    return render_template('contact.html')


if __name__ == '__main__':
    app.run(debug=True)

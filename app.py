import io
import os
import urllib.parse
from flask import Flask, redirect, render_template, request, send_file, url_for
from docx import Document

# १. सर्वात आधी app ची व्याख्या असावी:
app = Flask(__name__)

# २. मुख्य होमपेज
@app.route('/')
def home():
  return render_template('home.html')


# ३. दस्त मसुदा (मुख्य मेनू)
@app.route('/draft')
def draft():
  return render_template('draft_menu.html')


# ४. खरेदीखत उप-मेनू
@app.route('/kharedi_menu')
def kharedi_menu():
  return render_template('kharedi_menu.html')


# ५. प्लॉट खरेदीपत्र फॉर्म उघडणे
@app.route('/plot_draft')
def plot_draft():
  return render_template('plot_form.html')


# ६. सामान्य फॉर्म उघडणे
@app.route('/draft_form')
def draft_form():
  doc_type = request.args.get('type', '')
  return render_template('index.html', selected_type=doc_type)


# ७. प्लॉट खरेदीपत्र Word (.docx) डाऊनलोड करणे
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
  d = request.form
  doc = Document()

  p_title = doc.add_paragraph()
  p_title.alignment = 1
  p_title.add_run('|| श्री ||\n|| खरेदीपत्र ||\n').bold = True

  intro = (
      f"तुकडी व जिल्हा {d.get('taluka_dist')} मधील {d.get('village')} या गावाचे"
      ' नगरपालिका हद्दीतील गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व'
      ' त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये'
      f" {d.get('amount_num')} (अक्षरी रुपये {d.get('amount_words')})"
      f" प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {d.get('rate_sqm')} रुपये प्रमाणे दि."
      f" {d.get('doc_date')} इ.\n\n"
  )
  doc.add_paragraph(intro)

  doc.add_paragraph(
      f"{d.get('ghenar_name')}, {d.get('ghenar_age_occ')},\n"
      f"{d.get('ghenar_address')}\t\t\t\t\t\t\t\tलिहून घेणार\n"
      f"आ.क्र. {d.get('ghenar_aadhar')} पॅन नं. {d.get('ghenar_pan')}\n"
      f"मो.नं.:- {d.get('ghenar_mobile')}\n\n"
      'विरुद्ध\n\n'
      f"{d.get('denar_name')}, {d.get('denar_age_occ')},\n"
      f"{d.get('denar_address')}\t\t\t\t\t\t\t\tलिहून देणार\n"
      f"आ.क्र. {d.get('denar_aadhar')} पॅन नं. {d.get('denar_pan')}\n"
      f"मो.नं.:- {d.get('denar_mobile')}\n\n"
      'कारणे लिहून देतो खरेदीपत्र ऐसाजे की,\n'
  )

  milkat = (
      f"1) मिळकतीचे वर्णन:- \nजिल्हा व तुकडी कोल्हापूर पोट तुकडी"
      f" {d.get('taluka_dist')} मे. दुय्यम निबंधकसो यांचे अधिकार क्षेत्रातील"
      f" {d.get('village')} या गावाचे हद्दीतील मिळकत खरेदी देणार यांचे"
      ' स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे.\nरी.स.नं./गट'
      f" नं. {d.get('gat_no')} यासी एकूण क्षेत्र {d.get('total_area')} ची"
      ' मिळकत असून त्यापैकी खरेदी देणार यांचे मालकीचा गुंठेवारी नियमितीकरण'
      f" झालेले प्लॉट नंबर {d.get('plot_no')} याचे क्षेत्र {d.get('plot_area')}"
      f" ची खुली प्लॉट मिळकत व त्यामधील बांधकाम क्षेत्र {d.get('const_area')} ची"
      ' मिळकत सर्व.\n\nयांसी चतुःसिमा पुढीलप्रमाणे :-\nपूर्वेस :-'
      f" {d.get('east')} \nपश्चिमेस :- {d.get('west')} \nदक्षिणेस :-"
      f" {d.get('south')} \nउत्तरेस :- {d.get('north')} \n\nयेणेप्रमाणे"
      ' चतुःसिमेतील मिळकत सर्व व या मिळकतीमध्ये असलेली जल, तरू, काष्ट, पाषाण,'
      ' निधी-निक्षेप, तसेच जाणे-येणेच्या हक्कासह, आदी तदंगभूत वस्तूसहची'
      ' गुंठेवारी नियमितीकरण झालेली खुली मिळकत व त्यातील रो बंगलो मिळकत हि या'
      ' खरेदीपत्राचा विषय आहे.\n\n2) मोबदला व भरणा तपशील:-\nखरेदी मोबदल्याची'
      f" एकूण रक्कम रुपये {d.get('amount_num')} (अक्षरी रुपये"
      f" {d.get('amount_words')}) असून त्याचा तपशील"
      f" खालीलप्रमाणे:\n{d.get('payment_notes')}\n\n3) या खरेदीपत्राने"
      ' आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या हक्कासह सर्व मिळकतीचे'
      ' पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\nम्हणून स्वसंतोषाने व'
      ' अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता.'
      f" {d.get('doc_date')} इ."
  )
  doc.add_paragraph(milkat)

  target_stream = io.BytesIO()
  doc.save(target_stream)
  target_stream.seek(0)

  raw_filename = (
      f"PLOT खरेदीपत्र {d.get('village')} गट नं {d.get('gat_no')} प्लॉट नं"
      f" {d.get('plot_no')} {d.get('ghenar_name')}.docx"
  )
  clean_filename = raw_filename.replace('/', '-').replace('\\', '-')

  return send_file(
      target_stream,
      as_attachment=True,
      download_name=clean_filename,
      mimetype=(
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      ),
  )


# ८. IGR सब-पेज
@app.route('/igr')
def igr():
  return render_template('igr.html')


# ९. संपर्क पेज
@app.route('/contact')
def contact():
  return render_template('contact.html')

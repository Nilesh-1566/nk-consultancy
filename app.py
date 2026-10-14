import io
import os
import urllib.parse
from docx import Document
from flask import Flask, redirect, render_template, request, send_file, url_for

# १. Flask ॲप सुरू करणे
app = Flask(__name__)


# २. मुख्य होमपेज
@app.route('/')
def home():
  return render_template('home.html')


# ३. दस्त मसुदा मुख्य मेनू
@app.route('/draft')
def draft():
  return render_template('draft_menu.html')


# ४. खरेदीखत उप-मेनू (शेती, CTS, प्लॉट, करार TO खरेदीखत)
@app.route('/kharedi_menu')
def kharedi_menu():
  return render_template('kharedi_menu.html')


# ५. प्लॉट खरेदीपत्र डेटा नोंदणी फॉर्म
@app.route('/plot_draft')
def plot_draft():
  return render_template('plot_form.html')


# ६. सामान्य फॉर्म उघडणे
@app.route('/draft_form')
def draft_form():
  doc_type = request.args.get('type', '')
  return render_template('index.html', selected_type=doc_type)


# ७. प्लॉट खरेदीपत्र Word (.docx) तयार करून डाऊनलोड करणे
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
  form = request.form
  doc = Document()

  # दस्त शीर्षक
  p_title = doc.add_paragraph()
  p_title.alignment = 1
  p_title.add_run('|| श्री ||\n|| खरेदीपत्र ||\n').bold = True

  # प्रास्ताविक माहिती
  intro = (
      f"तुकडी व जिल्हा {form.get('taluka_dist')} मधील {form.get('village')} या"
      ' गावाचे नगरपालिका हद्दीतील गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट'
      ' जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये'
      f" {form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')})"
      f" प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {form.get('rate_sqm')} रुपये प्रमाणे"
      f" दि. {form.get('doc_date')} इ.\n\n"
  )
  doc.add_paragraph(intro)

  # १. लिहून घेणार (खरेदी घेणारा) तपशील
  g_names = form.getlist('ghenar_name[]')
  g_ages = form.getlist('ghenar_age_occ[]')
  g_addresses = form.getlist('ghenar_address[]')
  g_aadhars = form.getlist('ghenar_aadhar[]')
  g_pans = form.getlist('ghenar_pan[]')
  g_mobs = form.getlist('ghenar_mobile[]')

  p_ghenar = doc.add_paragraph()
  p_ghenar.add_run('लिहून घेणार (खरेदी घेणारा):\n').bold = True
  for i in range(len(g_names)):
    p_ghenar.add_run(
        f'{i+1}) {g_names[i]}, {g_ages[i] if i < len(g_ages) else ""}\n'
        f'{g_addresses[i] if i < len(g_addresses) else ""}\n'
        f'आ.क्र. {g_aadhars[i] if i < len(g_aadhars) else ""} | पॅन:'
        f' {g_pans[i] if i < len(g_pans) else ""} | मो.:'
        f' {g_mobs[i] if i < len(g_mobs) else ""}\n\n'
    )

  # २. लिहून देणार (खरेदी देणारा) तपशील
  d_names = form.getlist('denar_name[]')
  d_ages = form.getlist('denar_age_occ[]')
  d_addresses = form.getlist('denar_address[]')
  d_aadhars = form.getlist('denar_aadhar[]')
  d_pans = form.getlist('denar_pan[]')
  d_mobs = form.getlist('denar_mobile[]')

  p_denar = doc.add_paragraph()
  p_denar.add_run('विरुद्ध\n\nलिहून देणार (खरेदी देणारा):\n').bold = True
  for i in range(len(d_names)):
    p_denar.add_run(
        f'{i+1}) {d_names[i]}, {d_ages[i] if i < len(d_ages) else ""}\n'
        f'{d_addresses[i] if i < len(d_addresses) else ""}\n'
        f'आ.क्र. {d_aadhars[i] if i < len(d_aadhars) else ""} | पॅन:'
        f' {d_pans[i] if i < len(d_pans) else ""} | मो.:'
        f' {d_mobs[i] if i < len(d_mobs) else ""}\n\n'
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
    p_samati.add_run('संमती / मान्यता देणार:\n').bold = True
    for i in range(len(s_names)):
      if s_names[i].strip():
        p_samati.add_run(
            f'{i+1}) {s_names[i]}, {s_ages[i] if i < len(s_ages) else ""}\n'
            f'{s_addresses[i] if i < len(s_addresses) else ""}\n'
            f'आ.क्र. {s_aadhars[i] if i < len(s_aadhars) else ""} | पॅन:'
            f' {s_pans[i] if i < len(s_pans) else ""} | मो.:'
            f' {s_mobs[i] if i < len(s_mobs) else ""}\n\n'
        )

  # ४. मिळकतीचे वर्णन व कायदेशीर मजकूर
  milkat = (
      f'कारणे लिहून देतो खरेदीपत्र ऐसाजे की,\n'
      f'1) मिळकतीचे वर्णन:- \n'
      f'जिल्हा व तुकडी कोल्हापूर पोट तुकडी {form.get("taluka_dist")} मे. दुय्यम'
      ' निबंधकसो यांचे अधिकार क्षेत्रातील'
      f' {form.get("village")} या गावाचे हद्दीतील मिळकत खरेदी देणार यांचे'
      ' स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे.\n'
      f'री.स.नं./गट नं. {form.get("gat_no")} यामधील खरेदी देणार यांचे मालकीचा'
      ' गुंठेवारी नियमितीकरण झालेले प्लॉट नंबर'
      f' {form.get("plot_no")} याचे क्षेत्र {form.get("final_area")} ची खुली'
      ' प्लॉट मिळकत व त्यामधील बांधकाम क्षेत्र'
      f' {form.get("const_area")} ची मिळकत सर्व.\n\n'
      f'यांसी चतुःसिमा पुढीलप्रमाणे :-\n'
      f'पूर्वेस :- {form.get("east")} \n'
      f'पश्चिमेस :- {form.get("west")} \n'
      f'दक्षिणेस :- {form.get("south")} \n'
      f'उत्तरेस :- {form.get("north")} \n\n'
      'येणेप्रमाणे चतुःसिमेतील मिळकत सर्व व या मिळकतीमध्ये असलेली जल, तरू,'
      ' काष्ट, पाषाण, निधी-निक्षेप, तसेच जाणे-येणेच्या हक्कासह, आदी तदंगभूत'
      ' वस्तूसहची गुंठेवारी नियमितीकरण झालेली खुली मिळकत व त्यातील रो बंगलो'
      ' मिळकत हि या खरेदीपत्राचा विषय आहे.\n\n'
      f'2) मोबदला व भरणा तपशील:-\n'
      'खरेदी मोबदल्याची एकूण रक्कम रुपये'
      f' {form.get("amount_num")} (अक्षरी रुपये {form.get("amount_words")})'
      ' असून त्याचा तपशील खालीलप्रमाणे:\n'
  )
  doc.add_paragraph(milkat)

  # ५. भरणा तपशील यादी
  pay_amts = form.getlist('pay_amt[]')
  pay_modes = form.getlist('pay_mode[]')
  pay_banks = form.getlist('pay_bank[]')
  p_pay = doc.add_paragraph()
  for i in range(len(pay_amts)):
    if pay_amts[i].strip():
      p_pay.add_run(
          f'({i+1}) रक्कम रुपये {pay_amts[i]} -'
          f' {pay_modes[i] if i < len(pay_modes) else ""} द्वारे'
          f' {pay_banks[i] if i < len(pay_banks) else ""}. भरणा पावती निराळी'
          ' नाही.\n'
      )

  doc.add_paragraph(
      '\n3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या'
      ' हक्कासह सर्व मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला'
      ' आहात.\nम्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र'
      f' असे सही ता. {form.get("doc_date")} इ.'
  )

  # फाईल मेमरीमध्ये साठवून डाऊनलोड करणे
  target_stream = io.BytesIO()
  doc.save(target_stream)
  target_stream.seek(0)

  # मोबाईलवर डाऊनलोड फेल होऊ नये म्हणून सुरक्षित नाव (ASCII Safe Name)
  gat_no_clean = (
      str(form.get('gat_no', '87')).replace('/', '_').replace(' ', '_')
  )
  plot_no_clean = (
      str(form.get('plot_no', '26')).replace('/', '_').replace(' ', '_')
  )
  first_ghenar = (
      str(g_names[0]).strip().replace(' ', '_') if g_names else 'Ghenar'
  )

  filename = f'PLOT_Kharedipatra_Gat_{gat_no_clean}_Plot_{plot_no_clean}_{first_ghenar}.docx'

  return send_file(
      target_stream,
      as_attachment=True,
      download_name=filename,
      mimetype=(
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      ),
  )


# ८. IGR महाराष्ट्र माहिती पेज
@app.route('/igr')
def igr():
  return render_template('igr.html')


# ९. संपर्क माहिती पेज
@app.route('/contact')
def contact():
  return render_template('contact.html')


if __name__ == '__main__':
  app.run(debug=True)

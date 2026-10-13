from flask import Flask, render_template, request, send_file
from docx import Document
import io
import urllib.parse

# १. खरेदीखत मेनू
@app.route('/kharedi_menu')
def kharedi_menu():
    return render_template('kharedi_menu.html')

# २. प्लॉट खरेदीपत्र फॉर्म उघडणे
@app.route('/plot_draft')
def plot_draft():
    return render_template('plot_form.html')

# ३. भरलेल्या माहितीवरून अचूक Word (.docx) तयार करून डाउनलोड करणे
@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
    d = request.form

    doc = Document()

    # शीर्षक
    p_title = doc.add_paragraph()
    p_title.alignment = 1
    p_title.add_run("|| श्री ||\n|| खरेदीपत्र ||\n").bold = True

    # मुख्य माहिती परिच्छेद
    intro = (
        f"तुकडी व जिल्हा {d.get('taluka_dist')} मधील {d.get('village')} या गावाचे नगरपालिका हद्दीतील "
        f"गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये "
        f"{d.get('amount_num')} (अक्षरी रुपये {d.get('amount_words')}) प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {d.get('rate_sqm')} "
        f"रुपये प्रमाणे दि. {d.get('doc_date')} इ.\n\n"
    )
    doc.add_paragraph(intro)

    # पक्षकार तपशील
    doc.add_paragraph(
        f"{d.get('ghenar_name')}, {d.get('ghenar_age_occ')},\n"
        f"{d.get('ghenar_address')}\t\t\t\t\t\t\t\tलिहून घेणार\n"
        f"आ.क्र. {d.get('ghenar_aadhar')} पॅन नं. {d.get('ghenar_pan')}\n"
        f"मो.नं.:- {d.get('ghenar_mobile')}\n\n"
        f"विरुद्ध\n\n"
        f"{d.get('denar_name')}, {d.get('denar_age_occ')},\n"
        f"{d.get('denar_address')}\t\t\t\t\t\t\t\tलिहून देणार\n"
        f"आ.क्र. {d.get('denar_aadhar')} पॅन नं. {d.get('denar_pan')}\n"
        f"मो.नं.:- {d.get('denar_mobile')}\n\n"
        f"कारणे लिहून देतो खरेदीपत्र ऐसाजे की,\n"
    )

    # मिळकतीचे वर्णन
    milkat = (
        f"1) मिळकतीचे वर्णन:- \n"
        f"जिल्हा व तुकडी कोल्हापूर पोट तुकडी {d.get('taluka_dist')} मे. दुय्यम निबंधकसो यांचे अधिकार क्षेत्रातील "
        f"{d.get('village')} या गावाचे हद्दीतील मिळकत खरेदी देणार यांचे स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे.\n"
        f"री.स.नं./गट नं. {d.get('gat_no')} यासी एकूण क्षेत्र {d.get('total_area')} ची मिळकत असून त्यापैकी खरेदी देणार यांचे "
        f"मालकीचा गुंठेवारी नियमितीकरण झालेले प्लॉट नंबर {d.get('plot_no')} याचे क्षेत्र {d.get('plot_area')} ची खुली प्लॉट मिळकत "
        f"व त्यामधील बांधकाम क्षेत्र {d.get('const_area')} ची मिळकत सर्व.\n\n"
        f"यांसी चतुःसिमा पुढीलप्रमाणे :-\n"
        f"पूर्वेस :- {d.get('east')} \n"
        f"पश्चिमेस :- {d.get('west')} \n"
        f"दक्षिणेस :- {d.get('south')} \n"
        f"उत्तरेस :- {d.get('north')} \n\n"
        f"येणेप्रमाणे चतुःसिमेतील मिळकत सर्व व या मिळकतीमध्ये असलेली जल, तरू, काष्ट, पाषाण, निधी-निक्षेप, तसेच जाणे-येणेच्या हक्कासह, "
        f"आदी तदंगभूत वस्तूसहची गुंठेवारी नियमितीकरण झालेली खुली मिळकत व त्यातील रो बंगलो मिळकत हि या खरेदीपत्राचा विषय आहे.\n\n"
        f"2) मोबदला व भरणा तपशील:-\n"
        f"खरेदी मोबदल्याची एकूण रक्कम रुपये {d.get('amount_num')} (अक्षरी रुपये {d.get('amount_words')}) असून त्याचा तपशील खालीलप्रमाणे:\n"
        f"{d.get('payment_notes')}\n\n"
        f"3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या हक्कासह सर्व मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\n"
        f"म्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता. {d.get('doc_date')} इ."
    )
    doc.add_paragraph(milkat)

    # फाइलबफर तयार करणे
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    # तुम्ही सांगितल्याप्रमाणे नेमके फाईलचे नाव:
    # PLOT खरेदीपत्र त्यानंतर गावाचे नाव गट नंबर प्लॉट नंबर व लिहून घेणार यांचे नाव
    raw_filename = f"PLOT खरेदीपत्र {d.get('village')} गट नं {d.get('gat_no')} प्लॉट नं {d.get('plot_no')} {d.get('ghenar_name')}.docx"
    clean_filename = raw_filename.replace("/", "-").replace("\\", "-")

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=clean_filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

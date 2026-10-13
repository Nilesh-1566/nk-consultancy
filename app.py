@app.route('/generate_plot_docx', methods=['POST'])
def generate_plot_docx():
    form = request.form

    doc = Document()

    # शीर्षक
    p_title = doc.add_paragraph()
    p_title.alignment = 1
    p_title.add_run("|| श्री ||\n|| खरेदीपत्र ||\n").bold = True

    # मिळकतीची प्राथमिक माहिती
    intro = (
        f"तुकडी व जिल्हा {form.get('taluka_dist')} मधील {form.get('village')} या गावाचे नगरपालिका हद्दीतील "
        f"गुंठेवारी नियमितीकरण झालेल्या खुल्या प्लॉट जागेचे व त्यामधील बदिस्थ आर.सी.सी. इमारतीचे मिळकतीचे रक्कम रुपये "
        f"{form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) प्र.मु.वि.क्र. 19.1 प्रती चौ.मी. {form.get('rate_sqm')} "
        f"रुपये प्रमाणे दि. {form.get('doc_date')} इ.\n\n"
    )
    doc.add_paragraph(intro)

    # १. लिहून घेणार यादी
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
            f"रा. {g_addresses[i] if i < len(g_addresses) else ''}\n"
            f"आ.क्र. {g_aadhars[i] if i < len(g_aadhars) else ''} | पॅन: {g_pans[i] if i < len(g_pans) else ''} | मो.: {g_mobs[i] if i < len(g_mobs) else ''}\n\n"
        )

    # २. लिहून देणार यादी
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
            f"रा. {d_addresses[i] if i < len(d_addresses) else ''}\n"
            f"आ.क्र. {d_aadhars[i] if i < len(d_aadhars) else ''} | पॅन: {d_pans[i] if i < len(d_pans) else ''} | मो.: {d_mobs[i] if i < len(d_mobs) else ''}\n\n"
        )

    # ३. संमती देणार (असल्यास)
    s_names = form.getlist('samati_name[]')
    if s_names:
        p_samati = doc.add_paragraph()
        p_samati.add_run("संमती / मान्यता देणार:\n").bold = True
        for i in range(len(s_names)):
            p_samati.add_run(f"{i+1}) {s_names[i]}\n")
        p_samati.add_run("\n")

    # मिळकतीचे वर्णन
    milkat = (
        f"कारणे लिहून देतो खरेदीपत्र ऐसाजे की,\n"
        f"1) मिळकतीचे वर्णन:-\n"
        f"जिल्हा व तुकडी कोल्हापूर पोट तुकडी {form.get('taluka_dist')} मे. दुय्यम निबंधकसो यांचे अधिकार क्षेत्रातील "
        f"{form.get('village')} या गावाचे हद्दीतील मिळकत खरेदी देणार यांचे स्वसंपादित मालकी हक्काची व प्रत्यक्ष कब्जे वहीवाटीची आहे.\n"
        f"री.स.नं./गट नं. {form.get('gat_no')} यामधील खरेदी देणार यांचे मालकीचा गुंठेवारी नियमितीकरण झालेले प्लॉट नंबर {form.get('plot_no')} "
        f"याचे क्षेत्र {form.get('final_area')} ची खुली प्लॉट मिळकत व त्यामधील बांधकाम क्षेत्र {form.get('const_area')} ची मिळकत सर्व.\n\n"
        f"यांसी चतुःसिमा पुढीलप्रमाणे :-\n"
        f"पूर्वेस :- {form.get('east')} \n"
        f"पश्चिमेस :- {form.get('west')} \n"
        f"दक्षिणेस :- {form.get('south')} \n"
        f"उत्तरेस :- {form.get('north')} \n\n"
        f"येणेप्रमाणे चतुःसिमेतील मिळकत सर्व व या मिळकतीमध्ये असलेली जल, तरू, काष्ट, पाषाण, निधी-निक्षेप, तसेच जाणे-येणेच्या हक्कासह "
        f"गुंठेवारी नियमितीकरण झालेली खुली मिळकत व त्यातील रो बंगलो मिळकत हि या खरेदीपत्राचा विषय आहे.\n\n"
        f"2) मोबदला व भरणा तपशील:-\n"
        f"खरेदी मोबदल्याची एकूण रक्कम रुपये {form.get('amount_num')} (अक्षरी रुपये {form.get('amount_words')}) असून त्याचा तपशील खालीलप्रमाणे:\n"
    )
    doc.add_paragraph(milkat)

    # भरणा तपशील तक्ता किंवा यादी
    pay_amts = form.getlist('pay_amt[]')
    pay_modes = form.getlist('pay_mode[]')
    pay_banks = form.getlist('pay_bank[]')
    p_pay = doc.add_paragraph()
    for i in range(len(pay_amts)):
        p_pay.add_run(f"({i+1}) रक्कम रुपये {pay_amts[i]} - {pay_modes[i]} द्वारे {pay_banks[i]}. भरणा पावती निराळी नाही.\n")

    doc.add_paragraph(
        f"\n3) या खरेदीपत्राने आजपासून खरेदी घेणार यांनी वर वर्णन केलेल्या मिळकतीचे पूर्णपणे मालक व कब्जेदार वहिवाटदार झाला आहात.\n"
        f"म्हणून स्वसंतोषाने व अक्कल हुशारीने लिहून दिले कायम खूष खरेदीपत्र असे सही ता. {form.get('doc_date')} इ."
    )

    # फाईल तयार करणे व डाऊनलोड
    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)

    first_ghenar = g_names[0] if g_names else "खरेदीदार"
    raw_filename = f"PLOT खरेदीपत्र {form.get('village')} गट नं {form.get('gat_no')} प्लॉट नं {form.get('plot_no')} {first_ghenar}.docx"
    clean_filename = raw_filename.replace("/", "-").replace("\\", "-")

    return send_file(
        target_stream,
        as_attachment=True,
        download_name=clean_filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

@app.route('/')
def home():
    return render_template('home.html')

# १. 'दस्त ड्राफ्ट' बटन दाबल्यावर सर्व प्रकारांचे बटने दिसणारे पेज उघडेल
@app.route('/draft')
def draft():
    return render_template('draft_menu.html')

# २. त्या बटनावर क्लिक केल्यावर मूळ फॉर्म (index.html) उघडेल
@app.route('/draft_form')
def draft_form():
    doc_type = request.args.get('type', '')
    return render_template('index.html', selected_type=doc_type)

@app.route('/igr')
def igr():
    return render_template('igr.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

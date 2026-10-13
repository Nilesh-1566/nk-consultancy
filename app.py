@app.route('/')
def home():
    return render_template('home.html')

@app.route('/draft')
def draft():
    return render_template('index.html')

@app.route('/igr')
def igr():
    return render_template('igr.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

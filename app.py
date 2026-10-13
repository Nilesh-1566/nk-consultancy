from flask import Flask, render_template, request, send_file
# इतर आवश्यक imports...

app = Flask(__name__)

# १. होमपेज (ज्यावर सर्व बटने दिसतील)
@app.route('/')
def home():
    return render_template('home.html')

# २. दस्त ड्राफ्ट बटन दाबल्यावर उघडणारे पेज (तुमचे आधीचे मुख्य पेज)
@app.route('/draft')
def draft():
    return render_template('index.html')

# ३. संपर्क पेज
@app.route('/contact')
def contact():
    return render_template('contact.html')

# ... तुमचे बाकीचे जुने कोड जसेच्या तसे राहू द्या

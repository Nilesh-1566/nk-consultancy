from flask import Flask, render_template, request, send_file, redirect, url_for
import os

app = Flask(__name__)

# १. मुख्य होमपेज
@app.route('/')
def home():
    return render_template('home.html')

# २. दस्त ड्राफ्ट पेज
@app.route('/draft')
def draft():
    return render_template('index.html')

# ३. IGR महाराष्ट्र सब-पेज
@app.route('/igr')
def igr():
    return render_template('igr.html')

# ४. संपर्क माहिती पेज
@app.route('/contact')
def contact():
    return render_template('contact.html')

# याखाली तुमचा जुना दस्त तयार करण्याचा / टोकनचा बाकी सर्व कोड जसा आहे तसाच राहू द्या

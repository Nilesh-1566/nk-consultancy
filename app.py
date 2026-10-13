from flask import Flask, render_template, request, send_file, redirect, url_for
import os

app = Flask(__name__)

# १. मुख्य होमपेज
@app.route('/')
def home():
    return render_template('home.html')

# २. दस्त मसुदा (सर्व प्रकारांची बटने)
@app.route('/draft')
def draft():
    return render_template('draft_menu.html')

# ३. ठराविक दस्त निवडल्यावर उघडणारा फॉर्म
@app.route('/draft_form')
def draft_form():
    doc_type = request.args.get('type', '')
    return render_template('index.html', selected_type=doc_type)

# ४. IGR महाराष्ट्र सब-पेज
@app.route('/igr')
def igr():
    return render_template('igr.html')

# ५. संपर्क माहिती पेज
@app.route('/contact')
def contact():
    return render_template('contact.html')

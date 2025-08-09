from flask import Flask, render_template, request
import openpyxl
import pandas as pd
import os
from datetime import datetime
import requests  

app = Flask(__name__)

EXCEL_FILE = 'od-details.xlsx'

# Create Excel if not exists
if not os.path.exists(EXCEL_FILE):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "OD Entries"
    sheet.append(["Name", "Section", "Dept", "Roll", "Reason", "Time", "Submitted At"])
    workbook.save(EXCEL_FILE)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    try:
        data = {
            'name': request.form['name'],
            'section': request.form['section'],
            'dept': request.form['dept'],
            'roll': request.form['roll'],
            'reason': request.form['reason'],
            'time': request.form['time']
        }

        # Your Google Apps Script Web App URL
        script_url = 'https://script.google.com/macros/s/AKfycbxWrPgNhS1Moso56GqYjDAYapH5118ZeJ54jqXHyns-GnwgXLG_AzZlBwwYo0lsoj5mOw/exec'

        response = requests.post(script_url, json=data)

        if response.status_code == 200:
            resp_text = response.text.strip().lower()

            if resp_text == "duplicate":
                return "⚠ Duplicate entry detected! This Roll and Time already exist."
            elif resp_text == "success":
                return "✅ Form submitted successfully!"
            else:
                return f"Unexpected response: {response.text}"
        else:
            return f"Failed to submit: {response.text}"

    except Exception as e:
        return f"An error occurred: {e}"

if __name__ == '__main__':
    app.run(debug=True)


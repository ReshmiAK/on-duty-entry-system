from flask import Flask, render_template, request, redirect
import openpyxl
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

EXCEL_FILE = 'od-details.xlsx'

# 🔁 Rewrite Excel header without "Class"
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
        name = request.form['name']
        section = request.form['section']
        dept = request.form['dept']
        roll = request.form['roll']
        reason = request.form['reason']
        time = request.form['time']
        submission_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
        else:
            df = pd.DataFrame(columns=["Name", "Section", "Dept", "Roll", "Reason", "Time", "Submitted At"])

        new_row = pd.DataFrame([{
            "Name": name,
            "Section": section,
            "Dept": dept,
            "Roll": roll,
            "Reason": reason,
            "Time": time,
            "Submitted At": submission_time
        }])

        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)

        return "Form submitted successfully!"
    except Exception as e:
        return f"An error occurred: {e}"

if __name__ == '__main__':
    app.run(debug=True)

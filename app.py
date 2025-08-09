from flask import Flask, render_template, request
import os
import requests  

app = Flask(__name__)

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

        # Google Script Web App URL
        script_url = 'https://script.google.com/macros/s/AKfycbxWrPgNhS1Moso56GqYjDAYapH5118ZeJ54jqXHyns-GnwgXLG_AzZlBwwYo0lsoj5mOw/exec'

        # Step 1: Check for duplicate before submitting
        check_url = f"{script_url}?checkDuplicate=true&roll={data['roll']}&time={data['time']}"
        check_response = requests.get(check_url)

        if check_response.status_code == 200 and check_response.text.strip().lower() == "duplicate":
            return "Duplicate entry detected! Submission blocked."

        # Step 2: Submit if no duplicate
        response = requests.post(script_url, json=data)

        if response.status_code == 200:
            return "Form submitted successfully!"
        else:
            return f"Failed to submit: {response.text}"

    except Exception as e:
        return f"An error occurred: {e}"

if __name__ == '__main__':
    app.run(debug=True)

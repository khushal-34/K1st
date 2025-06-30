
from flask import Flask, render_template, request, redirect, send_file, session, url_for
import openpyxl
import os

app = Flask(__name__)
app.secret_key = 'supersecurekey'

EXCEL_FILE = 'data/client_data.xlsx'
PASSWORD = 'admin123'

# Ensure Excel file exists
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Name', 'Email', 'Message'])
    wb.save(EXCEL_FILE)

@app.route('/')
def client_form():
    return render_template('client.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    email = request.form['email']
    msg = request.form['message']

    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    ws.append([name, email, msg])
    wb.save(EXCEL_FILE)

    return render_template('client.html', submitted=True, name=name)

@app.route('/host', methods=['GET', 'POST'])
def host_login():
    if request.method == 'POST':
        if request.form['password'] == PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('host_panel'))
        else:
            return render_template('login.html', error="Wrong password")
    return render_template('login.html')

@app.route('/panel')
def host_panel():
    if not session.get('authenticated'):
        return redirect(url_for('host_login'))
    return render_template('host.html')

@app.route('/download')
def download():
    if not session.get('authenticated'):
        return redirect(url_for('host_login'))
    return send_file(EXCEL_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

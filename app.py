from flask import Flask, render_template, request, redirect, url_for
import openpyxl

app = Flask(__name__)

@app.route('/')
def form():
    return render_template('form.html') 

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form.get('name')
    email = request.form.get('email')
    date = request.form.get('date')
    checkin = request.form.get('checkin')
    checkout = request.form.get('checkout')
    task = request.form.get('task')
    work_report = request.form.get('work_report')
    remarks = request.form.get('remarks')

    try:
        workbook = openpyxl.load_workbook("employee_data.xlsx")
        sheet = workbook.active
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Name", "Email", "Date", "Check-in", "Check-out", "Task", "Work Report", "Remarks"])

    sheet.append([name, email, date, checkin, checkout, task, work_report, remarks])

    workbook.save("employee_data.xlsx")

    return redirect(url_for('success'))

@app.route('/success')
def success():
    return render_template('success.html') 

if __name__ == '__main__':
    app.run(debug=True)

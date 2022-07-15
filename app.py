from flask import *  
import os
import time

app = Flask(__name__)  
 
@app.route('/') 
def upload():  
    return render_template("index.html", logs='Upload Student USN and Subject Codes')  
 
@app.route('/upload_usns', methods = ['POST'])  
def usnsFile():  
    if request.method == 'POST':  
        f = request.files['file']  
        os.chdir('D:\web_scrap\input')
        f.save(f.filename)

        try:
            os.rename(f.filename, 'student_usn.csv')
        except WindowsError:
            os.remove('student_usn.csv')
            os.rename(f.filename, 'student_usn.csv')  
        return render_template("index.html", logs = 'USN upload successful')   

@app.route('/upload_codes', methods = ['POST'])  
def codesFile():  
    if request.method == 'POST':  
        f = request.files['file']  
        os.chdir('D:\web_scrap\input')
        f.save(f.filename)

        try:
            os.rename(f.filename, 'codes.csv')
        except WindowsError:
            os.remove('codes.csv')
            os.rename(f.filename, 'codes.csv')
        return render_template("index.html", logs='Codes upload successful') 

@app.route('/run', methods = ['POST'])
def run_script():
    py_filepath = 'D:\web_scrap\\vtu_result.py'
    status = os.system(f'py {py_filepath}')
    if status == 1:
        return render_template("index.html", logs='Result Extraction Successful')
    else:
        return render_template("index.html", logs='Wait for extraction to complete') 


@app.route('/download', methods = ['GET'])
def plot_csv():
    try:
        return send_file(
            'result\\marks.csv',
            mimetype='text/csv',
            download_name='marks.csv',
            as_attachment=True
        )
    except WindowsError:
        return render_template("index.html", logs='Results not found')  

if __name__ == '__main__':  
    app.run(debug = True)  
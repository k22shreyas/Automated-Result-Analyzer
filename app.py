import subprocess
import time
from flask import *  
import os
import re

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
        return render_template("index.html", logs='Subject Codes upload successful') 

@app.route('/run', methods = ['POST'])
def run_script():

    p = subprocess.Popen(['python','vtu_result.py'], stdout=subprocess.PIPE, bufsize=1)

    for line in iter(p.stdout.readline, b''):
        output = line.strip().decode( "utf-8" )
        print(output)

        if output == 'extractComplete':
            return render_template("index.html", logs='Result Extraction Successful')
        elif output == 'siteDown':
            return render_template("index.html", logs='VTU Site Down')
        elif output == 'noDriver':
            return render_template("index.html", logs='Chrome Driver')
        elif output == 'noWindow':
            return render_template("index.html", logs='Chrome Browser Closed') 
        else:
            return render_template("index.html", logs='Wait for extraction to complete...')
    p.stdout.close()
    p.wait()
    
@app.route('/download', methods = ['GET'])
def plot_csv():
    try:
        return send_file(
            'result\\student_marks.xlsx',
            mimetype='text/csv',
            download_name='student_marks.xlsx',
            as_attachment=True
        )
    except WindowsError:
        return render_template("index.html", logs='Results not found')  

if __name__ == '__main__':  
    app.run(debug = True)  
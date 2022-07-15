from flask import *  
import os
import subprocess
import sys
app = Flask(__name__)  
 
@app.route('/') 
def upload():  
    return render_template("index.html", downloadAlert='Wait for extraction to complete')  
 
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
        return render_template("index.html", name = 'USNs uploaded successfully', downloadAlert='Wait for extraction to complete')   

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
        return render_template("index.html", downloadAlert='Wait for extraction to complete') 

@app.route('/run', methods = ['POST'])
def run_script():
    py_filepath = 'D:\web_scrap\\vtu_result.py'
    os.system(f'py {py_filepath}')

@app.route('/download',methods = ['GET'])
def plot_csv():
    try:
        return send_file(
            'result\\marks.csv',
            mimetype='text/csv',
            download_name='marks.csv',
            as_attachment=True
        )
    except WindowsError:
        return render_template("index.html", downloadAlert='Results not found')  



if __name__ == '__main__':  
    app.run(debug = True)  
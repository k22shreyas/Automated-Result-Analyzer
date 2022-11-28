import subprocess
from flask import *  
import os

app = Flask(__name__)  
def restart():
    os.system("python app.py")

@app.route('/') 
def upload():  
    return render_template("index.html", logs='Upload Student USN and Subject Codes')  
 
@app.route('/upload_usns', methods = ['POST'])  
def usnsFile():  
    if request.method == 'POST':  
        f = request.files['file']  
        os.chdir('..\input')
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
        os.chdir('..')
        os.chdir('..\input')
        f.save(f.filename)

        try:
            os.rename(f.filename, 'codes.csv')
        except WindowsError:
            os.remove('codes.csv')
            os.rename(f.filename, 'codes.csv')
        return render_template("index.html", logs='Subject Codes upload successful') 

@app.route('/upload_link', methods = ['POST'])  
def linkFile():  
    if request.method == 'POST':  
        f = request.files['file']  
        os.chdir('..')
        os.chdir('..\input')
        f.save(f.filename)

        try:
            os.rename(f.filename, 'link.txt')
        except WindowsError:
            os.remove('link.txt')
            os.rename(f.filename, 'link.txt')
        return render_template("index.html", logs='Result Link upload successful') 

@app.route('/run', methods = ['POST'])
def run_script():
    try:
        os.chdir('..')
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

    except FileNotFoundError:
        return render_template("index.html", logs='Files not found')
    except TypeError:
        restart()
        return render_template("index.html", logs='Try Again...')

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

@app.route('/download_csv', methods = ['GET'])
def down_csv():
    try:
        return send_file(
            'result\\marks.csv',
            mimetype='text/csv',
            download_name='student_marks.csv',
            as_attachment=True,
        )
    except WindowsError:
        return render_template("index.html", logs='Results not found') 

if __name__ == '__main__':  
    app.run(debug = True)  

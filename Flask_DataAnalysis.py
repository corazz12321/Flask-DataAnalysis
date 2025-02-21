import os
import pandas as pd
from flask import Flask, request, redirect, url_for, render_template
from werkzeug.utils import secure_filename  # 修正 secure_filename 的導入
import json
from json2html import json2html

# 設定檔案上傳路徑（使用環境變數）
UPLOAD_FOLDER = os.path.join(os.environ["USERPROFILE"], "Desktop", "Flask_Upload")
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# 確保 UPLOAD_FOLDER 存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


# 檢查檔案副檔名是否允許
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/upload", methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        if 'file' not in request.files:
            return "No file part", 400
        
        file = request.files['file']
        
        if file.filename == '':
            return "No selected file", 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            return redirect(url_for('analyze'))
    
    return render_template('gui_data.html')


@app.route("/analyze")
def analyze():
    # 確保有檔案可以分析
    files = os.listdir(UPLOAD_FOLDER)
    if not files:
        return "No files found in the upload directory", 404

    for file in files:
        file_path = os.path.join(UPLOAD_FOLDER, file)
        
        try:
            # 處理 CSV 檔案
            if file.endswith(".csv"):
                df = pd.read_csv(file_path, sep=",")
                data = df.describe().to_html()
            # 處理 Excel 檔案
            elif file.endswith((".xlsx", ".xls")):
                df = pd.read_excel(file_path)  # `sep=","` 不適用於 Excel
                data = df.describe().to_html()
            else:
                continue  # 如果不是 CSV 或 Excel，跳過
            
            # 確保檔案已關閉後刪除
            os.remove(file_path)

            return data  # 只分析第一個符合的檔案，然後返回結果
        except Exception as e:
            return f"Error processing file {file}: {str(e)}", 500

    return "No valid CSV or Excel files found", 404


if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000, debug=True)

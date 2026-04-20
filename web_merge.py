from flask import Flask, request, send_file
import os
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return '''
    <h2>Excel 合併工具</h2>
    <form method="post" action="/merge" enctype="multipart/form-data">
        <input type="file" name="files" multiple>
        <br><br>
        <button type="submit">上傳並合併</button>
    </form>
    '''

@app.route("/merge", methods=["POST"])
def merge():
    files = request.files.getlist("files")
    dataframes = []

    for file in files:
        if file.filename.endswith(".xls") or file.filename.endswith(".xlsx"):
            filepath = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
            file.save(filepath)
            df = pd.read_excel(filepath)
            dataframes.append(df)

    if not dataframes:
        return "沒有有效的Excel檔案"

    merged = pd.concat(dataframes, ignore_index=True)
    output_path = os.path.join(UPLOAD_FOLDER, "merged.xlsx")
    merged.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

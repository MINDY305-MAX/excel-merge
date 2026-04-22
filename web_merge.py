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
    <h2>Excel 合併工具（完整版）</h2>
    <form method="post" action="/merge" enctype="multipart/form-data">
        <input type="file" name="files" multiple>
        <br><br>
        <button type="submit">上傳並合併</button>
    </form>
    '''

@app.route("/merge", methods=["POST"])
def merge():
    files = request.files.getlist("files")

    # 櫃號封條優先排序
    files = sorted(files, key=lambda x: ("櫃號封條" not in x.filename, x.filename))

    all_data = []
    errors = []

    for file in files:
        filename = secure_filename(file.filename)

        if not (filename.endswith(".xls") or filename.endswith(".xlsx")):
            errors.append(f"{filename} 不是Excel")
            continue

        try:
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            file.save(filepath)

            excel_file = pd.ExcelFile(filepath)

            # 櫃號封條 → 只抓第1個sheet
            if "櫃號封條" in filename:
                df = pd.read_excel(filepath, sheet_name=0, engine="xlrd")
                df["來源檔案"] = filename
                df["工作表"] = "Sheet1"
                all_data.append(df)
            else:
                # 其他 → 全部sheet
                for sheet in excel_file.sheet_names:
                    if filename.endswith(".xls"):
                        df = pd.read_excel(filepath, sheet_name=sheet, engine="xlrd")
                    else:
                        df = pd.read_excel(filepath, sheet_name=sheet)

                    df["來源檔案"] = filename
                    df["工作表"] = sheet
                    all_data.append(df)

        except Exception as e:
            errors.append(f"{filename} 錯誤: {str(e)}")

    if not all_data:
        return "全部檔案都讀取失敗<br>" + "<br>".join(errors)

    merged = pd.concat(all_data, ignore_index=True)

    output_path = os.path.join(UPLOAD_FOLDER, "merged.xlsx")
    merged.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)

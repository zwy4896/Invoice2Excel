from flask import Flask, render_template, request, jsonify
from werkzeug.utils import secure_filename
import os
from Invoice2Excel import Extractor
import pandas as pd

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'upload/'
ALLOWED_EXT = ('pdf')
OUT_PATH = 'test.xlsx'

# 判断文件后缀名
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT
    #  rsplit() 方法通过指定分隔符对字符串进行分割并返回一个列表

def run(file):
    extractor = Extractor(file)
    data = pd.DataFrame()
    frames = {}

    try:
        s = extractor.extract()
        s.name = os.path.basename(file)
        data = data.append(s)
        print(data)
    except Exception as e:
        print('file error:', file, '\n', e)
    frames[file.split('/')[-1]] = data
    with pd.ExcelWriter(OUT_PATH) as writer:
        for name, df in frames.items():
            df.to_excel(writer, sheet_name=name)
    print(f'{"*" * 50}\nALL DONE. THANK YOU FOR USING MY PROGRAMME. GOODBYE!\n{"*" * 50}')

@app.route('/', methods=['GET','POST'])
def upload_file():
    return render_template('upload.html')

@app.route('/upload', methods=['GET','POST'])
def uploader():
    if request.method == 'POST':
        f = request.files['file']
        if not (f and allowed_file(f.filename)):
            return jsonify({"error": 1001, "msg": "请检查上传的文件类型，仅限于pdf"})
        print(request.files)
        f.save(os.path.join(app.config['UPLOAD_FOLDER'],f.filename))
        run(os.path.join(app.config['UPLOAD_FOLDER'],f.filename))

        return 'file uploaded successfully'
    else:
        return render_template('upload.html')
if __name__ == '__main__':
   app.run(
        host='127.0.0.1',
        port=5000,
        debug=True
   )
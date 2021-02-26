import os
from flask import Flask, render_template, request, redirect, url_for, abort, send_file, after_this_request
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 1024 * 1024
app.config['UPLOAD_EXTENSIONS'] = ['.xls', '.xlsx']
app.config['UPLOAD_PATH'] = 'uploads'
app.config['OUTPUT_PATH'] = 'outputs'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods=['POST'])
def upload_files():
    if 'file' not in request.files:
        return render_template('result.html', message='식스샵에서 엑셀 다운받아서 올려주세요.')

    uploaded_file = request.files['file']
    filename = secure_filename(uploaded_file.filename)

    if filename != '':
        file_ext = os.path.splitext(filename)[1]
        if file_ext not in app.config['UPLOAD_EXTENSIONS']:
            return render_template('result.html', message='엑셀 파일이 아닙니다.')

    # 업로드 된 엑셀파일 저장
    filepath = os.path.join(app.config['UPLOAD_PATH'], filename)
    uploaded_file.save(filepath)

    # 엑셀 파일 처리
    converted_df = convert_order_format(filepath)

    # 엑셀 저장
    converted_filepath = os.path.join(app.config['OUTPUT_PATH'], "crayonbox-" + datetime.today().strftime('%Y-%m-%d') + '.xlsx') 
    converted_df.to_excel(converted_filepath, index=False)

    @after_this_request
    def remove_file(response):
        os.remove(filepath) # 업로드 된 엑셀 파일 삭제
        os.remove(converted_filepath) # 변환 된 엑셀 파일 삭제
        return response #render_template('result.html', message='정상 처리 완료!')

    return send_file(converted_filepath, as_attachment=True)

def convert_order_format (filepath):
    from_df = pd.read_excel(filepath)
    #print(from_df.head())

    to_df = pd.DataFrame(columns=['판매처', '상품명'])

    rows = []
    for index, item in from_df.iterrows():
        row = {}
        row['판매처'] = '자사몰'
        row['상품명'] = item['품목명']
        row['수량'] = item['주문 품목 개수']
        row['수령인명'] = item['배송지 이름']
        row['수령인주소'] = item['배송지 주소']
        row['우편번호'] = item['우편번호']
        row['핸드폰'] = item['배송지 휴대폰 번호']
        row['전화'] = item['배송지 휴대폰 번호']
        row['배송메시지'] = item['배송 시 요청 사항']

        rows.append(row)

    to_df = to_df.append(rows)

    return to_df

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8081, debug=True)
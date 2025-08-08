
from flask import Flask, request, jsonify
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Only Excel files are allowed'}), 400

    try:
        # Lưu file tạm thời vào thư mục hệ thống
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp:
            file.save(temp.name)
            temp_path = temp.name

        # Đọc dữ liệu bằng pandas
        df = pd.read_excel(temp_path).fillna("")

        # Lưu nguyên tên cột (không strip) để mapping
        fields_raw = [str(col) for col in df.columns]

        # Tạo header hiển thị/ghi TXT: thay thế ký tự đặc biệt, KHÔNG strip
        def sanitize_field(col):
            col = str(col)
            return col.replace('\t', '3==D').replace('\r\n', '3==D').replace('\n', '3==D').replace('\r', '3==D')
        fields_header = [sanitize_field(col) for col in df.columns]

        # Xử lý dữ liệu trong các ô - thay thế ký tự đặc biệt, KHÔNG đổi key
        def process_data_cell(value):
            if isinstance(value, str):
                return value.strip().replace('\t', '3==D').replace('\r\n', '3==D').replace('\n', '3==D').replace('\r', '3==D')
            return str(value)
        df_processed = df.applymap(process_data_cell)

        result = df_processed.to_dict(orient='records')

        # Xóa file tạm
        os.remove(temp_path)

        return jsonify({'fields_raw': fields_raw, 'fields': fields_header, 'data': result})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
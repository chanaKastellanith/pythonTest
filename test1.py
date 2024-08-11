from flask import Flask, request, jsonify
import pandas as pd
import os
from fpdf import FPDF
import matplotlib.pyplot as plt

app = Flask(__name__)


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and file.filename.endswith('.xlsx'):
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        xls = pd.ExcelFile(file_path)
        return jsonify({
            'file_path': file_path,
            'sheet_names': xls.sheet_names,
            'number_of_sheets': len(xls.sheet_names)
        })
    else:
        return jsonify({'error': 'Invalid file type'}), 400


from flask import Flask, request, jsonify
import pandas as pd

app = Flask(__name__)


@app.route('/process', methods=['POST'])
def process_data():
    data = request.json

    # בדוק אם המפתח 'operations' קיים במילון
    if 'operations' not in data:
        return jsonify({'error': 'Missing "operations" key in request data'}), 400

    operations = data['operations']

    file_path = data.get('file_path', '')
    if not file_path:
        return jsonify({'error': 'Missing "file_path" key in request data'}), 400

    report = {}
    xls = pd.ExcelFile(file_path)
    available_sheets = xls.sheet_names

    for operation_info in operations:
        sheet_name = operation_info.get('sheet_name')
        if sheet_name not in available_sheets:
            return jsonify({'error': f'Sheet {sheet_name} not found'}), 400

        operation = operation_info.get('operation')
        columns = operation_info.get('columns', [])

        df = pd.read_excel(file_path, sheet_name=sheet_name)

        # בדוק אם כל העמודות קיימות בגיליון
        missing_columns = [col for col in columns if col not in df.columns]
        if missing_columns:
            return jsonify({'error': f'Missing columns in sheet {sheet_name}: {missing_columns}'}), 400

        if operation == 'sum':
            result = df[columns].sum().to_dict()
        elif operation == 'average':
            result = df[columns].mean().to_dict()
        else:
            return jsonify({'error': 'Invalid operation'}), 400

        report[sheet_name] = result

    return jsonify(report)




@app.route('/generate_pdf', methods=['POST'])
def generate_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)

    for sheet_name, data in report.items():
        pdf.cell(200, 10, txt=sheet_name, ln=True)
        for col, val in data.items():
            pdf.cell(200, 10, txt=f"{col}: {val}", ln=True)

    pdf_output_path = 'report.pdf'
    pdf.output(pdf_output_path)

    return jsonify({'pdf_path': pdf_output_path})


@app.route('/plot', methods=['POST'])
def plot_graph():
    data = request.json
    sums = {sheet: sum(val.values()) for sheet, val in data.items()}
    sheets = list(sums.keys())
    values = list(sums.values())

    plt.bar(sheets, values)
    plt.xlabel('Sheet Names')
    plt.ylabel('Sum')
    plt.title('Sum of Each Sheet')
    plt.savefig('sheet_sums.png')

    return jsonify({'graph_path': 'sheet_sums.png'})


@app.route('/report', methods=['POST'])
def generate_report():
    data = request.json
    file_path = data['file_path']
    sheets_info = data['sheets']  # List of dicts with sheet name, operation, and columns

    report = {}
    for sheet_info in sheets_info:
        sheet_name = sheet_info['sheet_name']
        operation = sheet_info['operation']
        columns = sheet_info['columns']

        # Load the sheet
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        if operation == 'sum':
            result = df[columns].sum().to_dict()
        elif operation == 'average':
            result = df[columns].mean().to_dict()
        else:
            return jsonify({'error': 'Invalid operation'}), 400

        report[sheet_name] = result

    return jsonify(report)
@app.route('/generate_detailed_pdf', methods=['POST'])
def generate_detailed_pdf():
    report = request.json
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font('Arial', 'B', 16)

    pdf.cell(200, 10, txt="Report Details", ln=True)
    for sheet_name, data in report.items():
        pdf.cell(200, 10, txt=sheet_name, ln=True)
        for col, val in data.items():
            pdf.cell(200, 10, txt=f"{col}: {val}", ln=True)

    # Adding graphs
    pdf.add_page()
    pdf.cell(200, 10, txt="Graphs", ln=True)
    pdf.image('sheet_sums.png', x=10, y=30, w=180)

    pdf_output_path = 'detailed_report.pdf'
    pdf.output(pdf_output_path)

    return jsonify({'pdf_path': pdf_output_path})


if __name__ == '__main__':
    app.run(debug=True)

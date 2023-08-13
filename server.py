import os
import pandas as pd
import xlsxwriter
from flask import Flask, send_file, render_template, jsonify, request
from io import BytesIO
from datetime import datetime
import functions

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload_json', methods=['POST'])
def upload_json():
    try:
        json_file = request.files['jsonFile']
        if json_file and json_file.filename.endswith('.json'):
            json_filename = 'input.json'
            json_file.save(json_filename)
            response = {'json_file': json_filename}
            return jsonify(response), 200
        else:
            error_message = {'error': 'Invalid JSON file'}
            return jsonify(error_message), 400
    except Exception as e:
        error_message = {'error': str(e)}
        return jsonify(error_message), 500


@app.route('/convert_and_export', methods=['POST'])
def convert_and_export():
    try:
        json_filename = 'input.json'
        csv_filename = 'output.csv'
        excel_filename = 'output.xlsx'
        
        functions.convert_json_to_csv(json_filename, csv_filename)
        
        # Wait until the CSV file is ready before proceeding
        while not os.path.exists(csv_filename):
            time.sleep(2)  # Wait for 2 second
        functions.csv_to_excel(csv_filename, excel_filename)
        
        result = {
            'csv_file': csv_filename,
            'excel_file': excel_filename
        }
        
        return jsonify(result), 200  # Return HTTP 200 OK
    except FileNotFoundError:
        error_message = {'error': 'Input JSON file not found.'}
        return jsonify(error_message), 404  # Return HTTP 404 Not Found
    except ValueError as e:
        error_message = {'error': f'Invalid JSON data: {str(e)}'}
        return jsonify(error_message), 400  # Return HTTP 400 Bad Request
    except Exception as e:
        error_message = {'error': str(e)}
        return jsonify(error_message), 500  # Return HTTP 500 Internal Server Error



@app.route('/download_excel', methods=['GET'])
def download_excel():
    # Read CSV data into a pandas DataFrame
    csv_file = "output.csv"
    data = pd.read_csv(csv_file)

    # Create an Excel writer with XlsxWriter engine
    excel_output = BytesIO()  # Create an in-memory Excel file
    excel_writer = pd.ExcelWriter(excel_output, engine='xlsxwriter')

    # Write DataFrame to Excel
    data.to_excel(excel_writer, sheet_name="default", index=False)

    # Get the XlsxWriter workbook and worksheet objects
    workbook = excel_writer.book
    worksheet = excel_writer.sheets["default"]

    # Create a cell format with bold text and centered alignment
    bold_center_format = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter"})

    # Apply the cell format to the first row (header)
    worksheet.set_row(0, None, bold_center_format)

    # Set column widths
    column_widths = [max(len(str(value)) for value in data[column]) for column in data.columns]
    for i, width in enumerate(column_widths):
        worksheet.set_column(i, i, width + 2)  # Add some extra padding

    # Close the Excel writer (No need to call save())
    excel_writer.close()
    excel_output.seek(0)  # Reset the stream to the beginning

    # Serve the Excel file as a response
    return send_file(excel_output, as_attachment=True, download_name=f'excel_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')


@app.route('/clear_files', methods=['POST'])
def clear_files_route():
    project_directory = '/Your/Desktop/Project/Path/' # this will be relative to the 
    functions.clear_files(project_directory)
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
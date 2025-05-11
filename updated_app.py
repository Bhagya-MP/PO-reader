import os
from flask import Flask, request, jsonify, render_template
from utils import extract_data_from_pdf, load_product_dimensions, process_cargils_data, process_summary_order_data, process_country_style_data, process_Softlogic_data, process_Laugfs_data, process_Arpico_data, process_other_data

app = Flask(__name__)


@app.route('/upload', methods=['POST'])
def upload_pdf():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    file_path = os.path.join("uploads", file.filename)
    pdf_name = os.path.splitext(os.path.basename(file_path))[0] # Get the file name without extension for create excel file name based on pdf name
    os.makedirs("uploads", exist_ok=True)  # Ensure the upload folder exists
    file.save(file_path)

    try:
        extracted_data = extract_data_from_pdf(file_path)
        excel_path = "D:\DS_projects\OCR for PO\Diamension_table.xlsx"
        product_dimensions = load_product_dimensions(excel_path)

        # Additional logic for data extraction based on file name or content
        pdf_file_name = file.filename.lower()
        if "Cargills Summary".lower() in pdf_file_name:
            output_data = process_summary_order_data(extracted_data)
        elif "Cargills".lower() in pdf_file_name:
            output_data = process_cargils_data(extracted_data, product_dimensions)
        elif pdf_file_name == "country style.pdf":
            output_data = process_country_style_data(extracted_data, product_dimensions)
        elif "Softlogic".lower() in pdf_file_name:
            output_data = process_Softlogic_data(extracted_data, product_dimensions)
        elif "Laugfs P".lower() in pdf_file_name:
            output_data = process_Laugfs_data(extracted_data, product_dimensions)
        elif "Arpico".lower() in pdf_file_name:
            output_data = process_Arpico_data(extracted_data, product_dimensions)
        else:
            # Custom extraction logic for other PDFs
            output_data = process_other_data(extracted_data, product_dimensions)
            
        # create_excel(output_data, pdf_name)
        return jsonify(output_data), 200  # Return JSON response with status 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

if __name__ == '__main__':
    app.run(debug=True, port=5013)



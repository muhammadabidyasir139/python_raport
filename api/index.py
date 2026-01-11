from flask import Flask, request, send_file, jsonify
from docxtpl import DocxTemplate
import pandas as pd
import os
import zipfile
import io
import tempfile
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Konfigurasi
UPLOAD_FOLDER = tempfile.gettempdir()
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), '../templates')
DEFAULT_TEMPLATE_PATH = os.path.join(TEMPLATE_DIR, 'rapor_template.docx')

@app.route('/api/generate', methods=['POST'])
def generate_rapor():
    try:
        # 1. Cek File Excel
        if 'excel_file' not in request.files:
            return jsonify({"error": "No excel file provided"}), 400
        
        excel_file = request.files['excel_file']
        if excel_file.filename == '':
            return jsonify({"error": "No selected excel file"}), 400

        # 2. Cek Template (Upload atau Default)
        template_path = DEFAULT_TEMPLATE_PATH
        temp_template = None
        
        if 'template_file' in request.files and request.files['template_file'].filename != '':
            # Jika user upload template
            uploaded_tpl = request.files['template_file']
            temp_template = os.path.join(UPLOAD_FOLDER, secure_filename(uploaded_tpl.filename))
            uploaded_tpl.save(temp_template)
            template_path = temp_template
        elif not os.path.exists(template_path):
             return jsonify({"error": "Default template not found on server and no template uploaded"}), 500

        # 3. Proses Data
        df = pd.read_excel(excel_file)
        
        # Buffer untuk ZIP
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
            # Load template sekali
            doc = DocxTemplate(template_path)
            
            for index, row in df.iterrows():
                # Context data
                context = row.to_dict()
                
                # Render
                doc.render(context)
                
                # Nama file output
                nama = str(context.get('nama_lengkap', f'Siswa_{index}')).strip()
                kelas = str(context.get('Kelas', '')).strip()
                safe_name = "".join([c for c in nama if c.isalpha() or c.isdigit() or c==' ']).strip()
                safe_kelas = "".join([c for c in kelas if c.isalpha() or c.isdigit() or c==' ']).strip()
                
                filename = f"Rapor_{safe_name}_{safe_kelas}.docx"
                
                # Simpan ke temp stream
                file_stream = io.BytesIO()
                doc.save(file_stream)
                file_stream.seek(0)
                
                # Masukkan ke ZIP
                zip_file.writestr(filename, file_stream.getvalue())

        # Cleanup temp template jika ada
        if temp_template and os.path.exists(temp_template):
            os.remove(temp_template)

        zip_buffer.seek(0)
        
        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='rapor_generated.zip'
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

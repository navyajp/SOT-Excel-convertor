from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
import tempfile
import os
import traceback
import io

from parser import parse_nsdl_pdf, records_to_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.join(os.path.dirname(BASE_DIR), 'frontend')

app = Flask(__name__, static_folder=FRONTEND_DIR, static_url_path='')
CORS(app)

MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB


@app.route('/', methods=['GET'])
def index():
    return send_from_directory(FRONTEND_DIR, 'index.html')


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        records, client_name = parse_nsdl_pdf(tmp_path)

        if not records:
            return jsonify({'error': 'No transactions found in PDF. Please ensure this is a valid NSDL transaction statement.'}), 400

        excel_bytes = records_to_excel(records, client_name)
        safe_client_name = (client_name or 'output').replace(' ', '_')
        filename = f"NSDL_Transactions_{safe_client_name}.xlsx"

        return send_file(
            io.BytesIO(excel_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Processing failed: {str(e)}'}), 500
    finally:
        os.unlink(tmp_path)


@app.route('/preview', methods=['POST'])
def preview():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    file = request.files['file']
    if not file.filename or not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
        file.save(tmp.name)
        tmp_path = tmp.name

    try:
        records, client_name = parse_nsdl_pdf(tmp_path)

        return jsonify({
            'client_name': client_name,
            'total_records': len(records),
            'unique_securities': len(set(r['security_name'] for r in records)),
            'preview': records[:50]
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        os.unlink(tmp_path)


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_DEBUG', '').lower() in {'1', 'true', 'yes'}
    app.run(host='0.0.0.0', port=port, debug=debug)

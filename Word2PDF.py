from flask import Flask, render_template_string, request, send_file
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
import os
import sys
import uuid
import time
import re
import subprocess
import shutil
import tempfile
import logging

# ─────────────────────────────────────────────
# App setup
# ─────────────────────────────────────────────
app = Flask(__name__)

# FIX 1: File size limit — max 20 MB
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20 MB

# FIX 2: Rate limiting — max 10 requests per minute per IP
limiter = Limiter(
    get_remote_address,
    app=app,
    default_limits=["10 per minute"],
    storage_uri="memory://"
)

# FIX 3: Logging (replaces print statements)
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# HTML Template (unchanged from original)
# ─────────────────────────────────────────────
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<title>Word to PDF Converter</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
* { box-sizing: border-box; }
body { font-family: Arial, sans-serif; max-width: 600px; margin: 50px auto; padding: 20px; background: #f5f5f5; }
.container { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center; }
h1 { color: #333; margin-bottom: 10px; }
.subtitle { color: #666; margin-bottom: 30px; font-size: 14px; }
.upload-area { border: 2px dashed #4CAF50; padding: 40px; border-radius: 5px; margin: 20px 0; background: #fafafa; }
input[type="file"] { margin: 20px 0; padding: 10px; display: block; margin-left: auto; margin-right: auto; max-width: 100%; }
.btn { background: #4CAF50; color: white; padding: 12px 30px; border: none; border-radius: 5px; font-size: 16px; cursor: pointer; transition: background 0.3s; margin-top: 10px; }
.btn:hover { background: #45a049; }
.btn:disabled { background: #cccccc; cursor: not-allowed; }
#status { margin-top: 20px; color: #666; font-style: italic; min-height: 30px; }
.success { color: #4CAF50; font-weight: bold; }
.error { color: #f44336; font-weight: bold; }
.loading { color: #ff9800; font-weight: bold; }
#fileName { margin-top: 10px; color: #666; font-size: 14px; }
.features { margin-top: 30px; padding: 20px; background: #f9f9f9; border-radius: 5px; font-size: 13px; color: #666; }
.footer { margin-top: 30px; font-size: 12px; color: #999; }
@media (max-width: 480px) { .container { padding: 20px; } .upload-area { padding: 20px; } h1 { font-size: 24px; } }
</style>
</head>
<body>
<div class="container">
<h1>📄 Word to PDF Converter</h1>
<p class="subtitle">Upload a .docx file and convert it to PDF instantly!</p>
<div class="upload-area">
<form id="uploadForm" action="/convert" method="POST" enctype="multipart/form-data">
<input type="file" name="file" accept=".docx" required id="fileInput">
<br>
<button type="submit" class="btn" id="submitBtn">Convert to PDF</button>
</form>
<div id="fileName">No file selected</div>
</div>
<div id="status">Ready to convert...</div>
<div class="features">✅ Converts Word documents to PDF<br>✅ Preserves formatting<br>✅ Works on any device<br>✅ Max file size: 20MB</div>
<div class="footer">Made with ❤️ • Free & Open Source</div>
</div>
<script>
document.getElementById('fileInput').addEventListener('change', function(e) {
const fileName = e.target.files[0] ? e.target.files[0].name : 'No file selected';
document.getElementById('fileName').textContent = '📎 Selected: ' + fileName;
document.getElementById('status').textContent = 'Ready to convert...';
document.getElementById('status').className = '';
});
document.getElementById('uploadForm').onsubmit = async function(e) {
e.preventDefault();
const formData = new FormData(this);
const status = document.getElementById('status');
const submitBtn = document.getElementById('submitBtn');
const fileInput = document.getElementById('fileInput');
if (!fileInput.files || fileInput.files.length === 0) {
status.innerHTML = '❌ Please select a file first!';
status.className = 'error';
return;
}
status.innerHTML = '⏳ Converting... Please wait...';
status.className = 'loading';
submitBtn.disabled = true;
try {
const response = await fetch('/convert', { method: 'POST', body: formData });
if (response.ok) {
const contentDisposition = response.headers.get('Content-Disposition');
let filename = 'converted.pdf';
if (contentDisposition) {
const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
if (match && match[1]) { filename = match[1].replace(/['"]/g, ''); }
}
const blob = await response.blob();
const url = window.URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = filename;
document.body.appendChild(a);
a.click();
a.remove();
window.URL.revokeObjectURL(url);
status.innerHTML = '✅ Conversion successful! PDF downloaded as: ' + filename;
status.className = 'success';
} else {
const error = await response.text();
status.innerHTML = '❌ Error: ' + error;
status.className = 'error';
}
} catch (error) {
status.innerHTML = '❌ Error: ' + error.message;
status.className = 'error';
} finally {
submitBtn.disabled = false;
}
};
</script>
</body>
</html>
"""

# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def find_libreoffice():
    for cmd in ['libreoffice', 'soffice']:
        path = shutil.which(cmd)
        if path:
            return path
    return None


def convert_with_libreoffice(input_path, output_dir):
    """
    Converts a .docx file to PDF using LibreOffice.
    Returns the path to the generated PDF, or None on failure.
    """
    soffice_cmd = find_libreoffice()
    if not soffice_cmd:
        logger.error("LibreOffice not found on this system.")
        return None

    try:
        result = subprocess.run(
            [
                soffice_cmd,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                input_path
            ],
            capture_output=True,
            text=True,
            timeout=60
        )
        logger.info("LibreOffice stdout: %s", result.stdout.strip())
        if result.returncode != 0:
            logger.error("LibreOffice stderr: %s", result.stderr.strip())

        expected_pdf = os.path.join(
            output_dir,
            os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        )
        if os.path.exists(expected_pdf):
            return expected_pdf

        return None

    except subprocess.TimeoutExpired:
        logger.error("LibreOffice conversion timed out.")
        return None
    except Exception as e:
        logger.exception("LibreOffice error: %s", e)
        return None


# ─────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


# FIX 4: Rate limit applied specifically to /convert
@app.route('/convert', methods=['POST'])
@limiter.limit("10 per minute")
def convert_word_to_pdf():
    # FIX 5: Use OS temp directory instead of app root
    temp_dir = tempfile.mkdtemp()
    input_path = None
    output_path = None

    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400

        file = request.files['file']

        if not file or file.filename == '':
            return 'No file selected', 400

        # FIX 6: Validate file extension (case-insensitive)
        if not file.filename.lower().endswith('.docx'):
            return 'Only .docx files are supported', 400

        original_filename = file.filename
        name_without_ext = os.path.splitext(original_filename)[0]
        safe_name = sanitize_filename(name_without_ext)
        pdf_filename = f"{safe_name}.pdf"

        logger.info("Received file: %s", original_filename)

        unique_id = str(uuid.uuid4())
        input_path = os.path.join(temp_dir, f"input_{unique_id}.docx")

        file.save(input_path)
        logger.info("Saved input file: %s", input_path)

        # Convert
        generated_pdf = convert_with_libreoffice(input_path, temp_dir)

        if not generated_pdf or not os.path.exists(generated_pdf):
            return 'Conversion failed. LibreOffice may not be available.', 500

        # Rename to our expected output path
        output_path = os.path.join(temp_dir, f"output_{unique_id}.pdf")
        if generated_pdf != output_path:
            os.rename(generated_pdf, output_path)

        file_size = os.path.getsize(output_path)
        if file_size == 0:
            return 'Conversion failed — PDF is empty.', 500

        logger.info("PDF generated: %s (%d bytes)", output_path, file_size)

        response = send_file(
            output_path,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype='application/pdf'
        )

        # FIX 7: Reliable cleanup — delete entire temp_dir after response
        @response.call_on_close
        def cleanup():
            try:
                shutil.rmtree(temp_dir, ignore_errors=True)
                logger.info("Temp directory cleaned up: %s", temp_dir)
            except Exception as e:
                logger.warning("Cleanup error: %s", e)

        return response

    except Exception as e:
        # FIX 8: Generic error message to user — no stack trace exposed
        logger.exception("Unhandled error during conversion: %s", e)
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass
        return 'An internal error occurred. Please try again.', 500


# ─────────────────────────────────────────────
# Handle file-too-large error gracefully
# ─────────────────────────────────────────────
@app.errorhandler(413)
def file_too_large(e):
    return 'File too large. Maximum allowed size is 20 MB.', 413


# ─────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))

    logger.info("=" * 60)
    logger.info("Word to PDF Converter — Starting on port %d", port)
    logger.info("=" * 60)

    libreoffice = find_libreoffice()
    if libreoffice:
        logger.info("LibreOffice found at: %s", libreoffice)
    else:
        logger.warning("LibreOffice NOT found. Conversions will fail.")

    # FIX 1 (CRITICAL): debug=False in production
    app.run(host='0.0.0.0', port=port, debug=False)
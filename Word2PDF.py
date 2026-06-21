from flask import Flask, render_template_string, request, send_file
import os
import sys
import uuid
import time
import re
import subprocess
import shutil

app = Flask(__name__)

# [HTML TEMPLATE - SAME AS BEFORE]
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
        <div class="features">✅ Converts Word documents to PDF<br>✅ Preserves formatting<br>✅ Works on any device</div>
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

def sanitize_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def find_libreoffice():
    """Find LibreOffice executable"""
    # Try common names
    possible_cmds = ['libreoffice', 'soffice']
    for cmd in possible_cmds:
        path = shutil.which(cmd)
        if path:
            return path
    return None

def convert_with_libreoffice(input_path, output_path):
    """Convert using LibreOffice"""
    soffice_cmd = find_libreoffice()
    print(f"🔍 LibreOffice found at: {soffice_cmd}")
    
    if not soffice_cmd:
        return False
    
    try:
        # Run conversion
        result = subprocess.run([
            soffice_cmd,
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', os.path.dirname(output_path),
            input_path
        ], capture_output=True, text=True, timeout=60)
        
        # Check if PDF was created
        expected_pdf = os.path.join(
            os.path.dirname(output_path),
            os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
        )
        
        if os.path.exists(expected_pdf):
            if expected_pdf != output_path:
                os.rename(expected_pdf, output_path)
            return True
        
        return False
        
    except Exception as e:
        print(f"❌ LibreOffice error: {e}")
        return False

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert_word_to_pdf():
    input_path = None
    output_path = None
    
    try:
        if 'file' not in request.files:
            return 'No file uploaded', 400
        
        file = request.files['file']
        if file.filename == '':
            return 'No file selected', 400
        if not file.filename.lower().endswith('.docx'):
            return 'Only .docx files are supported', 400
        
        original_filename = file.filename
        name_without_extension = os.path.splitext(original_filename)[0]
        safe_name = sanitize_filename(name_without_extension)
        pdf_filename = f"{safe_name}.pdf"
        
        print(f"📄 Uploaded file: {original_filename}")
        print(f"💻 Platform: {sys.platform}")
        
        unique_id = str(uuid.uuid4())
        input_path = os.path.abspath(f'input_{unique_id}.docx')
        output_path = os.path.abspath(f'output_{unique_id}.pdf')
        
        file.save(input_path)
        print(f"📁 File saved: {input_path}")
        
        # Convert using LibreOffice
        success = convert_with_libreoffice(input_path, output_path)
        
        if not success:
            return 'Conversion failed. LibreOffice is not available.', 500
        
        if not os.path.exists(output_path):
            return 'Conversion failed - PDF not created', 500
        
        file_size = os.path.getsize(output_path)
        if file_size == 0:
            return 'Conversion failed - PDF is empty', 500
        
        print(f"📊 PDF size: {file_size} bytes")
        
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype='application/pdf'
        )
        
        @response.call_on_close
        def cleanup():
            try:
                if os.path.exists(input_path):
                    os.remove(input_path)
                if os.path.exists(output_path):
                    time.sleep(0.5)
                    os.remove(output_path)
            except Exception as e:
                print(f"⚠️ Cleanup error: {e}")
        
        return response
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return f'Error: {str(e)}', 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60)
    print("🚀 Word to PDF Converter - Starting...")
    print("=" * 60)
    print(f"📱 Server running on port: {port}")
    
    # Check for LibreOffice
    libreoffice = find_libreoffice()
    if libreoffice:
        print(f"✅ LibreOffice found at: {libreoffice}")
    else:
        print("⚠️ LibreOffice not found! Please ensure it's installed.")
    
    print("⏹️  Press CTRL+C to stop the server")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=True)
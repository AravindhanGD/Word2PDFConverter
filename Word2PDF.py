from flask import Flask, render_template_string, request, send_file
import os
from docx2pdf import convert
import uuid
import time
import re

app = Flask(__name__)

# HTML template for the webpage
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Word to PDF Converter</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 600px;
            margin: 50px auto;
            padding: 20px;
            background: #f5f5f5;
        }
        .container {
            background: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            text-align: center;
        }
        h1 {
            color: #333;
            margin-bottom: 30px;
        }
        .upload-area {
            border: 2px dashed #4CAF50;
            padding: 40px;
            border-radius: 5px;
            margin: 20px 0;
            background: #fafafa;
        }
        input[type="file"] {
            margin: 20px 0;
            padding: 10px;
            display: block;
            margin-left: auto;
            margin-right: auto;
        }
        .btn {
            background: #4CAF50;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            cursor: pointer;
            transition: background 0.3s;
            margin-top: 10px;
        }
        .btn:hover {
            background: #45a049;
        }
        .btn:disabled {
            background: #cccccc;
            cursor: not-allowed;
        }
        #status {
            margin-top: 20px;
            color: #666;
            font-style: italic;
            min-height: 30px;
        }
        .success {
            color: #4CAF50;
            font-weight: bold;
        }
        .error {
            color: #f44336;
            font-weight: bold;
        }
        .loading {
            color: #ff9800;
            font-weight: bold;
        }
        #fileName {
            margin-top: 10px;
            color: #666;
            font-size: 14px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 Word to PDF Converter</h1>
        <p>Upload a .docx file and convert it to PDF instantly!</p>
        
        <div class="upload-area">
            <form id="uploadForm" action="/convert" method="POST" enctype="multipart/form-data">
                <input type="file" name="file" accept=".docx" required id="fileInput">
                <br>
                <button type="submit" class="btn" id="submitBtn">Convert to PDF</button>
            </form>
            <div id="fileName">No file selected</div>
        </div>
        
        <div id="status">Ready to convert...</div>
    </div>

    <script>
        // Show selected filename
        document.getElementById('fileInput').addEventListener('change', function(e) {
            const fileName = e.target.files[0] ? e.target.files[0].name : 'No file selected';
            document.getElementById('fileName').textContent = '📎 Selected: ' + fileName;
            document.getElementById('status').textContent = 'Ready to convert...';
            document.getElementById('status').className = '';
        });

        // Handle form submission
        document.getElementById('uploadForm').onsubmit = async function(e) {
            e.preventDefault();
            
            const formData = new FormData(this);
            const status = document.getElementById('status');
            const submitBtn = document.getElementById('submitBtn');
            
            // Check if file is selected
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
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    // Get the filename from the response headers
                    const contentDisposition = response.headers.get('Content-Disposition');
                    let filename = 'converted.pdf';
                    if (contentDisposition) {
                        const match = contentDisposition.match(/filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/);
                        if (match && match[1]) {
                            filename = match[1].replace(/['"]/g, '');
                        }
                    }
                    
                    // Download the PDF
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
    """Remove invalid characters from filename"""
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return filename

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['GET', 'POST'])
def convert_word_to_pdf():
    # If someone visits /convert directly with GET, redirect to home
    if request.method == 'GET':
        return "Please use the form at the homepage to upload a file. <a href='/'>Go back</a>", 405
    
    input_path = None
    output_path = None
    
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            return 'No file uploaded. Please select a .docx file.', 400
        
        file = request.files['file']
        
        if file.filename == '':
            return 'No file selected. Please choose a .docx file.', 400
        
        if not file.filename.lower().endswith('.docx'):
            return 'Only .docx files are supported. Please upload a Word document.', 400
        
        # Get the original filename without extension
        original_filename = file.filename
        name_without_extension = os.path.splitext(original_filename)[0]
        safe_name = sanitize_filename(name_without_extension)
        pdf_filename = f"{safe_name}.pdf"
        
        print(f"📄 Uploaded file: {original_filename}")
        print(f"📄 PDF will be: {pdf_filename}")
        
        # Generate unique filenames for temporary storage
        unique_id = str(uuid.uuid4())
        input_path = os.path.abspath(f'input_{unique_id}.docx')
        output_path = os.path.abspath(f'output_{unique_id}.pdf')
        
        # Save uploaded file
        file.save(input_path)
        print(f"📁 File saved temporarily as: {input_path}")
        
        # Convert Word to PDF
        print(f"🔄 Converting to PDF... This may take a moment...")
        convert(input_path, output_path)
        print(f"✅ PDF created at: {output_path}")
        
        # Wait for the file to be fully written
        time.sleep(1.5)
        
        # Check if PDF was created successfully
        if not os.path.exists(output_path):
            return 'Conversion failed - PDF was not created. Please check if Microsoft Word is installed correctly.', 500
        
        # Get file size to verify it's not empty
        file_size = os.path.getsize(output_path)
        if file_size == 0:
            return 'Conversion failed - PDF file is empty. The Word document might be corrupted.', 500
        
        print(f"📊 PDF size: {file_size} bytes")
        
        # Send the PDF file with the original filename
        response = send_file(
            output_path,
            as_attachment=True,
            download_name=pdf_filename,
            mimetype='application/pdf'
        )
        
        # Clean up files after response is sent
        @response.call_on_close
        def cleanup():
            try:
                if input_path and os.path.exists(input_path):
                    os.remove(input_path)
                    print(f"🧹 Cleaned up: {input_path}")
                if output_path and os.path.exists(output_path):
                    time.sleep(0.5)
                    os.remove(output_path)
                    print(f"🧹 Cleaned up: {output_path}")
            except Exception as e:
                print(f"⚠️ Cleanup error: {e}")
        
        return response
                
    except Exception as e:
        print(f"❌ Error during conversion: {str(e)}")
        # Clean up if error occurs
        try:
            if input_path and os.path.exists(input_path):
                os.remove(input_path)
            if output_path and os.path.exists(output_path):
                try:
                    time.sleep(0.5)
                    os.remove(output_path)
                except:
                    pass
        except:
            pass
        return f"Conversion error: {str(e)}", 500

if __name__ == '__main__':
    # Get the port from environment variable (Render sets this automatically)
    # If not set, default to 5000 for local testing
    port = int(os.environ.get('PORT', 5000))
    
    print("=" * 60)
    print("🚀 Word to PDF Converter - Starting...")
    print("=" * 60)
    print(f"📱 Server running on port: {port}")
    print("📄 Upload any .docx file and it will be converted to PDF")
    print("⏹️  Press CTRL+C to stop the server")
    print("=" * 60)
    
    # host='0.0.0.0' makes it accessible from anywhere (needed for cloud)
    app.run(host='0.0.0.0', port=port)
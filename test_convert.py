from docx2pdf import convert

# Test conversion
input_file = "test.docx"  # Create a simple test docx file
output_file = "test.pdf"

try:
    convert(input_file, output_file)
    print("✅ Conversion successful!")
except Exception as e:
    print(f"❌ Error: {e}")
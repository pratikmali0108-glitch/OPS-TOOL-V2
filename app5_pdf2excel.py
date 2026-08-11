from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pdfplumber
import pandas as pd
import os
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)
app.secret_key = 'your-secret-key-change-this-pdf2excel'  # Change this to a random secret key

# Configuration
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'pdf'}

# Create directories if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size for PDFs

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_excel(pdf_path):
    """
    Convert PDF to Excel using the same logic as the original script
    """
    try:
        all_rows = []
        page_info = []
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            print(f"Total pages in PDF: {total_pages}\n")
            
            for i, page in enumerate(pdf.pages):
                print(f"🔄 Extracting from page {i+1}...")
                page_info.append(f"Extracting from page {i+1}...")
                
                tables = page.extract_tables()
                
                if tables:
                    for table in tables:
                        for row in table:
                            # Clean each cell in the row
                            cleaned_row = [cell.strip() if cell else '' for cell in row]
                            # Add row only if it contains some data
                            if any(cleaned_row):
                                all_rows.append(cleaned_row)
                else:
                    # If no tables found, try to extract text line by line
                    text = page.extract_text()
                    if text:
                        lines = text.split('\n')
                        for line in lines:
                            if line.strip():
                                # Split line by multiple spaces or tabs to create columns
                                row = [cell.strip() for cell in line.split() if cell.strip()]
                                if row:
                                    all_rows.append(row)
        
        if not all_rows:
            return None, "No data could be extracted from the PDF. The PDF might be image-based or contain no tables/text."
        
        # Find maximum number of columns
        max_cols = max(len(row) for row in all_rows) if all_rows else 0
        
        if max_cols == 0:
            return None, "No valid data rows found in the PDF."
        
        # Normalize all rows to have the same number of columns
        normalized_rows = [row + [''] * (max_cols - len(row)) for row in all_rows]
        
        # Create DataFrame
        df = pd.DataFrame(normalized_rows)
        
        # Generate column names (A, B, C, etc.)
        column_names = []
        for i in range(max_cols):
            if i < 26:
                column_names.append(chr(65 + i))  # A, B, C, ..., Z
            else:
                # For columns beyond Z, use AA, AB, AC, etc.
                first_letter = chr(65 + (i // 26) - 1)
                second_letter = chr(65 + (i % 26))
                column_names.append(first_letter + second_letter)
        
        df.columns = column_names[:max_cols]
        
        extraction_info = {
            'total_pages': total_pages,
            'total_rows': len(all_rows),
            'total_columns': max_cols,
            'page_info': page_info
        }
        
        return df, extraction_info
        
    except Exception as e:
        return None, f"Error processing PDF: {str(e)}"

@app.route('/')
def index():
    return render_template('indexpdf2excel.html')

@app.route('/process', methods=['POST'])
def process_file():
    try:
        # Debug: Print all form data and files
        print("Form data:", dict(request.form))
        print("Files:", dict(request.files))
        
        # Check if file was uploaded
        if 'file' not in request.files:
            print("ERROR: 'file' not in request.files")
            flash('No file part in the request', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        
        print(f"File object: {file}")
        print(f"File filename: '{file.filename}'")
        
        # Validate inputs
        if not file or file.filename == '' or file.filename is None:
            print("ERROR: No file selected or empty filename")
            flash('Please select a PDF file to upload', 'error')
            return redirect(url_for('index'))
        
        if not allowed_file(file.filename):
            print(f"ERROR: Invalid file type: {file.filename}")
            flash('Invalid file type. Please upload PDF files only', 'error')
            return redirect(url_for('index'))
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        unique_filename = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        
        # Convert PDF to Excel
        result_df, extraction_info = convert_pdf_to_excel(file_path)
        
        if result_df is None:
            flash(f'Error processing PDF: {extraction_info}', 'error')
            # Clean up uploaded file
            if os.path.exists(file_path):
                os.remove(file_path)
            return redirect(url_for('index'))
        
        # Save processed file
        original_name = os.path.splitext(filename)[0]
        output_filename = f"{original_name}_converted_{timestamp}.xlsx"
        output_path = os.path.join(app.config['PROCESSED_FOLDER'], output_filename)
        result_df.to_excel(output_path, index=False)
        
        # Clean up uploaded file
        if os.path.exists(file_path):
            os.remove(file_path)
        
        # Create success message with extraction info
        if isinstance(extraction_info, dict):
            success_msg = f'PDF converted successfully! Extracted {extraction_info["total_rows"]} rows and {extraction_info["total_columns"]} columns from {extraction_info["total_pages"]} pages.'
        else:
            success_msg = 'PDF converted successfully!'
            
        flash(success_msg, 'success')
        return send_file(output_path, as_attachment=True, download_name=output_filename)
        
    except Exception as e:
        flash(f'An error occurred: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['PROCESSED_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            flash('File not found', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error downloading file: {str(e)}', 'error')
        return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
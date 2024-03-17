# Import necessary modules
from flask import Flask, render_template, request, redirect, flash
from werkzeug.utils import secure_filename
import os
import pytesseract
from PIL import Image
import pandas as pd
import re

# Create Flask application
app = Flask(__name__, template_folder='templates')
app.secret_key = 'RussLoveTheProg'
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


# Ensure the UPLOAD_FOLDER directory exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Function to preprocess image
def preprocess_image(image_path):
    image = Image.open(image_path)
    gray_image = image.convert("L")
    return gray_image

def perform_ocr(image_path, languages):
    image = Image.open(image_path)
    extracted_text = pytesseract.image_to_string(image, lang=languages)
    return extracted_text

def extract_contact_info(text):
    email = re.search(r'[\w\.-]+@[\w\.-]+', text).group() if re.search(r'[\w\.-]+@[\w\.-]+', text) else ''
    
    name_patterns = [
        r'([A-Z][a-z]+)\s([A-Z][a-z]+)',              # First and Last Name Separated by Space   
]

    
    # Iterate through patterns and try to match
    name = ''
    for pattern in name_patterns:
        name_match = re.search(pattern, text)
        if name_match:
            name = name_match.group(1).strip()
            break
    
    phone = re.search(r'Phone:\s*([^\n]+)', text).group(1) if re.search(r'Phone:\s*([^\n]+)', text) else ''
    
    company_match = re.search(r'@([\w\.-]+)', email)
    if company_match:
        company_name = company_match.group(1)
        website = f'www.{company_name}'
    else:
        company_name = ''
        website = ''

    return {'Name': name, 'Email': email, 'Phone': phone, 'Company': company_name, 'Website': website}



def process_images(file_paths, languages):
    data = []
    supported_formats = ['.jpg', '.jpeg', '.png']
    for file_path in file_paths:
        if any(file_path.lower().endswith(format) for format in supported_formats):
            extracted_text = perform_ocr(file_path, languages)
            contact_info = extract_contact_info(extracted_text)
            data.append(contact_info)
            
    return data

# Route for the upload page
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'GET':
        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)
        files = request.files.getlist('files[]')
        if len(files) == 0:
            flash('No selected files')
            return redirect(request.url)
        file_paths = []
        for file in files:
            if file.filename == '':
                flash('No selected file')
                return redirect(request.url)
            if file:
                filename = secure_filename(file.filename)
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)
                file_paths.append(file_path)
        if file_paths:
            languages = ['eng']  # English and Arabic
            processed_data = process_images(file_paths, languages)
            df = pd.DataFrame(processed_data)
            output_file = 'contact_info.xlsx'
            df.to_excel(output_file, index=False)
            flash(f"Contact information saved to {output_file}")
            return render_template('result.html', processed_data=processed_data)
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)

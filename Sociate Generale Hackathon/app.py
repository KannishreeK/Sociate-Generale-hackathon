from flask import Flask, request, render_template
import os
from oletools.olevba import VBA_Parser
import logging

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Set up logging
logging.basicConfig(level=logging.DEBUG)

def extract_vba_code(file_path):
    try:
        vba_modules = []

        # Parse the Excel file to extract VBA code
        vba_parser = VBA_Parser(file_path)
        
        # Iterate through each VBA module
        for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
            vba_modules.append({
                'module_name': vba_filename,
                'code': vba_code.strip()  # Remove leading/trailing whitespace
            })
        
        vba_parser.close()

        return vba_modules

    except Exception as e:
        logging.error(f"Error extracting VBA code: {str(e)}")
        return None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']
    if uploaded_file.filename != '':
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
        uploaded_file.save(file_path)
        
        if file_path.endswith('.xlsm'):
            vba_modules = extract_vba_code(file_path)
            logging.debug(f"Extracted VBA modules: {vba_modules}")

            if vba_modules:
                # Perform analysis on the extracted VBA modules (if needed)
                analyzed_data = {'modules': vba_modules}  # Placeholder for analysis

                if analyzed_data:
                    return render_template('results.html', data=analyzed_data)
                else:
                    return "No VBA code analyzed."
            else:
                return "No VBA code extracted."
        else:
            return "Uploaded file must be in .xlsm format."

    return 'No file uploaded'

if __name__ == '__main__':
    app.run(debug=True)

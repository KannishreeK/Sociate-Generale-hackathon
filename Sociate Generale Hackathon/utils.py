import zipfile
import xml.etree.ElementTree as ET
import logging
import win32com.client as win32
import re

def extract_vba_code(file_path):
    vba_code = {}
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            if 'xl/vbaProject.bin' in zip_ref.namelist():
                vba_project_file = zip_ref.read('xl/vbaProject.bin')
            else:
                logging.error(f"'xl/vbaProject.bin' not found in '{file_path}'")
                return None
    
        vba_root = ET.fromstring(vba_project_file)
    
        for sheet in vba_root.findall('.//sheet'):
            sheet_name = sheet.get('name')
            vba_code[sheet_name] = []
            for obj in sheet.findall('.//script'):
                vba_code[sheet_name].append(obj.text)
    
        return vba_code if vba_code else None
    except FileNotFoundError:
        logging.error(f"File '{file_path}' not found.")
        return None
    except zipfile.BadZipFile:
        logging.error(f"File '{file_path}' is not a valid zip file.")
        return None
    except ET.ParseError as e:
        logging.error(f"XML ParseError: {str(e)}")
        return None
    except Exception as e:
        logging.error(f"Error extracting VBA code: {str(e)}")
        return None

def add_vba_macro(file_path, macro_code):
    try:
        # Open Excel application
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        
        # Open the workbook
        workbook = excel.Workbooks.Open(file_path)
        
        # Check if VBA project is locked
        if workbook.VBProject.Protection == 1:
            print("The VBA project is locked. Unable to add macro.")
            workbook.Close(SaveChanges=False)
            excel.Quit()
            return False

        # Add a new module
        module = workbook.VBProject.VBComponents.Add(1)  # 1 represents a standard module
        module.CodeModule.AddFromString(macro_code)
        
        # Save and close the workbook
        workbook.Close(SaveChanges=True)
        excel.Quit()
        return True
    except Exception as e:
        print(f"Error adding VBA macro: {str(e)}")
        return False

def analyze_vba_code(vba_code):
    analyzed_data = {}

    function_pattern = re.compile(r'(Sub|Function)\s+(\w+)')
    variable_pattern = re.compile(r'\bDim\s+(\w+)')
    
    for sheet_name, code_lines in vba_code.items():
        analyzed_data[sheet_name] = {
            'functions': [],
            'variables': [],
        }

        for line in code_lines:
            function_match = function_pattern.search(line)
            if function_match:
                function_name = function_match.group(2)
                analyzed_data[sheet_name]['functions'].append(function_name)

            variable_match = variable_pattern.search(line)
            if variable_match:
                variable_name = variable_match.group(1)
                analyzed_data[sheet_name]['variables'].append(variable_name)

    return analyzed_data if analyzed_data else None

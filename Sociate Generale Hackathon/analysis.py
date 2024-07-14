import re

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

import os
import re
import shutil
import google.generativeai as genai 
def extract_javascript(text):
    if "`" in text:
        clean_query = text.replace("`", "").replace("javascript", "")
        return clean_query
    else:
        return text

# Function to convert VBScript code to JavaScript
# def convert_vbscript_to_javascript(vbscript):
#     # Conversion patterns
#         genai.configure(api_key='AIzaSyDE8a285hL5J0lR7B1pxWOdY5GTlb51pjw')
#         model = genai.GenerativeModel('gemini-1.5-flash-8b')  
#         prompt = f"""
#     You are a developer assistant helping convert VBScript code to JavaScript. Please carefully translate the following VBScript code into its JavaScript equivalent. ***Be as accurate and clear as possible without any comments or content as i am directly the pasting the generated code into a file:***

#     VBScript:
#     {vbscript}

#     JavaScript:
#     """
#         response = model.generate_content(f"You are a helpful assistant. User: {prompt}")
#         if response and hasattr(response, 'candidates') and len(response.candidates) > 0:
#             answer = response.candidates[0].content.parts[0].text
#             sql_response = extract_javascript(answer)
#             # print(sql_response)
#             return sql_response

"""
    Sairam's code
    """

def vbscript_to_javascript(vbscript_code):
    """
    Converts VBScript code to JavaScript code while accounting for corner cases.
    """
    # Replace `Dim` for variable declarations
    js_code = re.sub(r"\bDim\b", "let", vbscript_code)
 
    # Replace `Set` with direct assignment
    js_code = re.sub(r"\bSet\b\s+", "", js_code)
 
    # Replace `If ... Then ... Else ... End If` with JavaScript's `if-else`
    js_code = re.sub(r"If (.+?) Then", r"if (\1) {", js_code)
    js_code = re.sub(r"ElseIf (.+?) Then", r"} else if (\1) {", js_code)
    js_code = re.sub(r"\bElse\b", "} else {", js_code)
    js_code = re.sub(r"End If", "}", js_code)
 
    # Replace `For Each ... In ... Next` loops with JavaScript's `for...of`
    js_code = re.sub(
        r"For Each (.+?) In (.+?)", r"for (const \1 of \2) {", js_code
    )
    js_code = re.sub(r"Next", "}", js_code)
 
    # Replace `For ... To ... Step` loops with JavaScript's `for` loops
    js_code = re.sub(
        r"For (\w+) = (\d+) To (\d+)",
        lambda match: f"for (let {match.group(1)} = {match.group(2)}; "
                      f"{match.group(1)} <= {match.group(3)}; "
                      f"{match.group(1)}++) {{",
        js_code,
    )
    js_code = re.sub(r"Next", "}", js_code)
 
    # Replace `While ... Wend` loops with `while`
    js_code = re.sub(r"While (.+?)\s", r"while (\1) {", js_code)
    js_code = re.sub(r"Wend", "}", js_code)
 
    # Replace `Do ... Loop` constructs
    js_code = re.sub(r"Do While (.+?)", r"while (\1) {", js_code)
    js_code = re.sub(r"Loop", "}", js_code)
 
    # Replace `Select Case ... End Select` with `switch`
    js_code = re.sub(r"Select Case (.+?)", r"switch (\1) {", js_code)
    js_code = re.sub(r"Case (.+?):", r"case \1:", js_code)
    js_code = re.sub(r"Case Else", "default:", js_code)
    js_code = re.sub(r"End Select", "}", js_code)
 
    # Replace VBScript functions with JavaScript equivalents
    js_code = re.sub(r"\bCStr\((.+?)\)", r"String(\1)", js_code)
    js_code = re.sub(r"\bCLng\((.+?)\)", r"Number(\1)", js_code)
    js_code = re.sub(r"\bUCase\((.+?)\)", r"\1.toUpperCase()", js_code)
    js_code = re.sub(r"\bLCase\((.+?)\)", r"\1.toLowerCase()", js_code)
    js_code = re.sub(r"\bLen\((.+?)\)", r"\1.length", js_code)
    js_code = re.sub(r"\bMid\((.+?),(.+?),(.+?)\)", r"\1.substr(\2 - 1, \3)", js_code)
 
    # Replace `MsgBox` with `alert` in JavaScript
    js_code = re.sub(r"MsgBox\s*\((.+?)\)", r"alert(\1)", js_code)
 
    # Replace function definitions
    js_code = re.sub(r"Function (\w+)", r"function \1", js_code)
    js_code = re.sub(r"End Function", "}", js_code)
 
    # Replace `Sub ... End Sub` with JavaScript function
    js_code = re.sub(r"Sub (\w+)", r"function \1", js_code)
    js_code = re.sub(r"End Sub", "}", js_code)
 
    # Replace VBScript string concatenation (`&`) with JavaScript's `+`
    js_code = re.sub(r" & ", " + ", js_code)
 
    # Replace `On Error Resume Next` with try-catch blocks
    if "On Error Resume Next" in js_code:
        js_code = re.sub(
            r"On Error Resume Next",
            "// Error handling starts here\ntry {",
            js_code,
        )
        js_code += "\n} catch (e) {\n  console.error(e);\n}"
 
    # General cleanup: removing VBScript specific keywords
    js_code = re.sub(r"\bOption Explicit\b", "// JavaScript does not require this", js_code)
 
    return js_code    

"""
    End of sairam code
    """
# Function to process ASP file and extract VBScript to convert to JavaScript
def process_asp_file(asp_file_path, output_folder):
    with open(asp_file_path, 'r', encoding='utf-8-sig') as asp_file:
        content = asp_file.read()
    vbscript_pattern=r'<SCRIPT[^>]*>.*?</SCRIPT>'
    vbscript_blocks = re.findall(vbscript_pattern, content, re.DOTALL)

    for vbscript in vbscript_blocks:
        print("------------------------")
        print(asp_file_path)
        # print(vbscript)
        js_code = vbscript_to_javascript(vbscript)
        content = content.replace(vbscript, js_code)

    # Save the updated ASP file in the new folder structure
    relative_path = os.path.relpath(asp_file_path, root_directory)
    new_path = os.path.join(output_folder, relative_path)
    os.makedirs(os.path.dirname(new_path), exist_ok=True)

    with open(new_path, 'w', encoding='utf-8') as new_file:
        new_file.write(content)

# Function to process VBS file and convert to JavaScript
def process_vbs_file(vbs_file_path, output_folder):
    with open(vbs_file_path, 'r', encoding='utf-8-sig') as vbs_file:
        vbscript = vbs_file.read()

    # Convert VBScript to JavaScript
    js_code = vbscript_to_javascript(vbscript)

    # Save the updated JavaScript file in the new folder structure
    relative_path = os.path.relpath(vbs_file_path, root_directory)
    new_path = os.path.join(output_folder, relative_path.replace('.vbs', '.js'))
    os.makedirs(os.path.dirname(new_path), exist_ok=True)

    with open(new_path, 'w', encoding='utf-8') as new_file:
        new_file.write(js_code)

# Main function to traverse directories and process files
def process_directory(root_directory, output_directory):
    # Traverse the ASP directory for ASP files
    asp_folder = os.path.join(root_directory, 'ASP')
    for root, _, files in os.walk(asp_folder):
        for file in files:
            if file.endswith('.asp'):
                asp_file_path = os.path.join(root, file)
                process_asp_file(asp_file_path, output_directory)

    # Traverse the VBS directory for VBS files
    vbs_folder = os.path.join(root_directory, 'VBS')
    for root, _, files in os.walk(vbs_folder):
        for file in files:
            if file.endswith('.vbs'):
                vbs_file_path = os.path.join(root, file)
                process_vbs_file(vbs_file_path, output_directory)

if __name__ == "__main__":
    # Set the root directory and output directory
    root_directory = r'C:\Users\ShashankTudum\Downloads\AQS_VBScript\AQS_VBScript'  # Update with your root directory
    output_directory = r'C:\Users\ShashankTudum\Downloads\shashank-output'  # Update with your output directory

    # Ensure output directory exists
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Process the directories
    process_directory(root_directory, output_directory)

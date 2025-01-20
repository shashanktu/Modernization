import xmltodict
import json
def xml_to_json(xml_file, json_file):
    # Open the XML file and parse it
    with open(xml_file, 'r', encoding='utf-8') as xml:
        xml_data = xmltodict.parse(xml.read())
    
    # Convert the parsed XML data to JSON format
    with open(json_file, 'w', encoding='utf-8') as jsonf:
        json.dump(xml_data, jsonf, indent=4, ensure_ascii=False)
    
    print(f"XML data has been successfully converted to JSON and saved to {json_file}")

# Example usage
xml_to_json(r'C:\Users\ShashankTudum\Downloads\shashank\sample.xml', r'C:\Users\ShashankTudum\Downloads\shashank-output\output.json')

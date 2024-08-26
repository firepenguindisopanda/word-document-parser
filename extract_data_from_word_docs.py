import os
import json
from docx import Document
from lxml import etree
import unicodedata

heading_keywords = [
        "Explore Majors", "Sample Career", "Prospective", "Skills And Characteristics",
        "Alternative Career", "Employers", "Areas", "Places that hire", "What can I do",
        "Skills Related", "Sample Jobs", "Sample"
    ]

def clean_text(text):
    """
    Normalize and clean up Unicode characters.
    param text (str): The text to clean up.
    return (str): The cleaned-up text.
    """
    # Normalize the text to decompose characters (e.g., accented characters)
    normalized_text = unicodedata.normalize('NFKD', text)
    # Encode to ASCII bytes, ignoring non-ASCII characters, then decode back to string
    cleaned_text = normalized_text.encode('ascii', 'ignore').decode('ascii')
    
    return cleaned_text




def extract_headings_and_content(file_path, heading_keywords):
    '''
    Extracts the headings and content from a Word document.

    param file_path (str): The path to the Word document file.
    param heading_keywords (list): A list of keywords that indicate a heading.
    return (dict): A dictionary where the keys are headings and the values are lists of content under each heading.
    '''
    document_content = Document(file_path)

    # Extract the XML content of the document
    xml_content = etree.fromstring(document_content.element.xml)
    # Define namespaces
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    # Initialize an object to store the content under each heading
    content_by_heading = {}
    # Iterate through the XML elements to find headings and their content
    current_heading = None
    for elem in xml_content.xpath('//w:p', namespaces=namespaces):
        # Extract the text of the paragraph
        paragraph_text = ''.join(elem.xpath('.//w:t/text()', namespaces=namespaces)).strip()
        paragraph_text = clean_text(paragraph_text)
        # Check if the paragraph is a heading
        if any(keyword in paragraph_text for keyword in heading_keywords):
            # If a new heading is found, update the current heading
            current_heading = paragraph_text
            content_by_heading[current_heading] = []
        elif current_heading:
            # If content is found under the current heading, add it to the content list
            if elem.xpath('.//w:tbl', namespaces=namespaces):
                # If the content is a table, extract the table rows
                table_content = []
                for row in elem.xpath('.//w:tr', namespaces=namespaces):
                    row_content = []
                    for cell in row.xpath('.//w:tc', namespaces=namespaces):
                        cell_text = ''.join(cell.xpath('.//w:t/text()', namespaces=namespaces)).strip()
                        cell_text = clean_text(cell_text)
                        if cell_text:
                            row_content.append(cell_text)
                    if row_content:
                        table_content.append(row_content)
                if table_content:
                    content_by_heading[current_heading].append({'table': table_content})
            else:
                # If the content is a paragraph, bullet point, or list, extract the text
                if paragraph_text:
                    content_by_heading[current_heading].append(paragraph_text)
    return content_by_heading

def process_documents(file_list_path):
    '''
    Process a list of Word documents to extract headings and content.
    param file_list_path (str): The path to the file containing a list of Word document paths.
    return (dict): A dictionary where the keys are document names and the values are the extracted headings and content.
    '''
    all_stats = {}

    with open(file_list_path, 'r') as file_list:
        for line in file_list:
            file_path = line.strip()
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                extracted_data_raw = extract_headings_and_content(file_path, heading_keywords=heading_keywords)
                all_stats[file_name] = extracted_data_raw
            else:
                print(f"Warning: File '{file_path}' not found.")

    return all_stats


def save_to_json(data, json_file_path):
    '''
    Save extracted data to a JSON file.
    param data (dict): The extracted data to save.
    param json_file_path (str): The path to the JSON file to save the data.
    '''
    with open(json_file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)

if __name__ == "__main__":
    file_list_path = "valid_docx_paths.txt"
    output_json_path = "extracted_careers_data.json"

    document_stats = process_documents(file_list_path)
    save_to_json(document_stats, output_json_path)

    print(f"Document raw data saved to {output_json_path}")

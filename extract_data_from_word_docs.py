import os
import json
from docx import Document
from lxml import etree
import unicodedata
import re

heading_keywords = [
        "Explore Majors", "Sample Career", "Prospective", "Skills And Characteristics",
        "Alternative Career", "Employers", "Areas", "Places that hire", "What can I do",
        "Skills Related", "Sample Jobs", "Sample"
    ]

def clean_text(text):
    """
    Normalize and clean up text by:
    - Converting to lowercase
    - Removing extra whitespace
    - Normalizing Unicode characters
    - Removing non-ASCII characters
    - Standardizing quotation marks and apostrophes
    
    Args:
        text (str): The text to clean up.
    Returns:
        str: The cleaned-up text.
    """
    if not text or not isinstance(text, str):
        return ""
        
    # Convert to lowercase
    text = text.lower()
    
    # Normalize Unicode characters
    text = unicodedata.normalize('NFKD', text)
    
    # Remove non-ASCII characters while preserving spaces
    text = ''.join(char if ord(char) < 128 else ' ' for char in text)
    
    # Standardize quotes and apostrophes
    text = re.sub(r'[''`]', "'", text)
    text = re.sub(r'["""]', '"', text)
    
    # Replace multiple spaces, tabs, and newlines with a single space
    text = re.sub(r'\s+', ' ', text)
    
    # Strip leading and trailing whitespace
    text = text.strip()
    
    return text

def extract_table_content(table_elem, namespaces):
    """
    Extract and clean content from a table element.
    
    Args:
        table_elem: The XML table element
        namespaces: XML namespaces dictionary
    Returns:
        list: List of cleaned table rows
    """
    table_content = []
    for row in table_elem.xpath('.//w:tr', namespaces=namespaces):
        row_content = []
        for cell in row.xpath('.//w:tc', namespaces=namespaces):
            cell_text = ''.join(cell.xpath('.//w:t/text()', namespaces=namespaces))
            cell_text = clean_text(cell_text)
            if cell_text:
                row_content.append(cell_text)
        if row_content:
            table_content.append(row_content)
    return table_content

def extract_headings_and_content(file_path, heading_keywords):
    """
    Extracts the headings and content from a Word document with improved text cleaning.

    Args:
        file_path (str): The path to the Word document file.
        heading_keywords (list): A list of keywords that indicate a heading.
    Returns:
        dict: A dictionary where the keys are headings and the values are lists of content.
    """
    try:
        document_content = Document(file_path)
    except Exception as e:
        print(f"Error reading document {file_path}: {str(e)}")
        return {}

    # Extract the XML content of the document
    try:
        xml_content = etree.fromstring(document_content.element.xml)
    except Exception as e:
        print(f"Error parsing XML in document {file_path}: {str(e)}")
        return {}

    # Define namespaces
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Initialize content storage
    content_by_heading = {}
    current_heading = None

    # Compile regex patterns for heading detection
    heading_patterns = [re.compile(rf'\b{re.escape(keyword)}\b', re.IGNORECASE) 
                       for keyword in heading_keywords]

    for elem in xml_content.xpath('//w:p', namespaces=namespaces):
        # Extract and clean paragraph text
        paragraph_text = ''.join(elem.xpath('.//w:t/text()', namespaces=namespaces))
        paragraph_text = clean_text(paragraph_text)
        
        # Skip empty paragraphs
        if not paragraph_text:
            continue

        # Check if paragraph is a heading
        is_heading = any(pattern.search(paragraph_text) for pattern in heading_patterns)
        
        if is_heading:
            current_heading = paragraph_text
            content_by_heading[current_heading] = []
        elif current_heading:
            # Handle table content
            tables = elem.xpath('.//w:tbl', namespaces=namespaces)
            if tables:
                for table in tables:
                    table_content = extract_table_content(table, namespaces)
                    if table_content:
                        content_by_heading[current_heading].append({'table': table_content})
            # Handle regular paragraph content
            elif paragraph_text:
                content_by_heading[current_heading].append(paragraph_text)

    return content_by_heading

def process_documents(file_list_path):
    """
    Process a list of Word documents to extract headings and content.
    
    Args:
        file_list_path (str): Path to the file containing Word document paths.
    Returns:
        dict: Dictionary of document names and their extracted content.
    """
    all_stats = {}

    try:
        with open(file_list_path, 'r') as file_list:
            for line in file_list:
                file_path = line.strip()
                if not file_path:
                    continue
                    
                if os.path.exists(file_path):
                    file_name = os.path.basename(file_path)
                    extracted_data = extract_headings_and_content(
                        file_path, 
                        heading_keywords=heading_keywords
                    )
                    if extracted_data:
                        all_stats[file_name] = extracted_data
                else:
                    print(f"Warning: File '{file_path}' not found.")
    except Exception as e:
        print(f"Error processing document list: {str(e)}")
        return {}

    return all_stats


def save_to_json(data, json_file_path):
    """
    Save extracted data to a JSON file with error handling.
    
    Args:
        data (dict): The extracted data to save.
        json_file_path (str): The path to save the JSON file.
    """
    try:
        with open(json_file_path, 'w', encoding='utf-8') as json_file:
            json.dump(data, json_file, indent=4, ensure_ascii=False)
        print(f"Successfully saved data to {json_file_path}")
    except Exception as e:
        print(f"Error saving JSON file: {str(e)}")

if __name__ == "__main__":
    file_list_path = "valid_docx_paths.txt"
    output_json_path = "extracted_careers_data.json"

    document_stats = process_documents(file_list_path)
    if document_stats:
        save_to_json(document_stats, output_json_path)

import os
import json
import yaml
from docx import Document
from lxml import etree
import unicodedata
import re
from pathlib import Path

def clean_text(text):
    """Normalize and clean up text."""
    if not text or not isinstance(text, str):
        return ""
    text = text.lower()
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(char if ord(char) < 128 else ' ' for char in text)
    text = re.sub(r'[''`]', "'", text)
    text = re.sub(r'["""]', '"', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    return text

def extract_table_content(table_elem, namespaces):
    """Extract and clean content from a table element."""
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
    """Extracts headings and content from a Word document."""
    try:
        document_content = Document(file_path)
        xml_content = etree.fromstring(document_content.element.xml)
    except Exception as e:
        print(f"Error reading/parsing document {file_path}: {str(e)}")
        return {}

    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    content_by_heading = {}
    current_heading = None

    heading_patterns = [re.compile(rf'\b{re.escape(keyword)}\b', re.IGNORECASE) 
                       for keyword in heading_keywords]

    for elem in xml_content.xpath('//w:p', namespaces=namespaces):
        paragraph_text = ''.join(elem.xpath('.//w:t/text()', namespaces=namespaces))
        paragraph_text = clean_text(paragraph_text)
        
        if not paragraph_text:
            continue

        is_heading = any(pattern.search(paragraph_text) for pattern in heading_patterns)
        
        if is_heading:
            current_heading = paragraph_text
            content_by_heading[current_heading] = []
        elif current_heading:
            tables = elem.xpath('.//w:tbl', namespaces=namespaces)
            if tables:
                for table in tables:
                    table_content = extract_table_content(table, namespaces)
                    if table_content:
                        content_by_heading[current_heading].append({'table': table_content})
            elif paragraph_text:
                content_by_heading[current_heading].append(paragraph_text)

    return content_by_heading

def run_parser(base_path: Path, manifest_path: str, output_path: str, heading_keywords: list):
    """Main entry point for parsers."""
    full_manifest_path = base_path / manifest_path
    all_data = {}

    if not full_manifest_path.exists():
        print(f"Error: Manifest {full_manifest_path} not found.")
        return False

    with open(full_manifest_path, 'r') as file_list:
        for line in file_list:
            rel_path = line.strip()
            if not rel_path: continue
            
            file_path = base_path / rel_path
            if file_path.exists():
                extracted_data = extract_headings_and_content(file_path, heading_keywords)
                if extracted_data:
                    all_data[file_path.name] = extracted_data
            else:
                print(f"Warning: File '{file_path}' not found.")

    with open(base_path / output_path, 'w', encoding='utf-8') as f:
        json.dump(all_data, f, indent=4, ensure_ascii=False)
    
    print(f"Successfully saved parsed data to {output_path}")
    return True

import os
import json
import unicodedata
from docx import Document
from pathlib import Path

def clean_text(text):
    normalized_text = unicodedata.normalize('NFC', text)
    cleaned_text = ' '.join(normalized_text.split())
    return cleaned_text

def is_heading(text, heading_keywords):
    return any(text.lower().startswith(keyword.lower()) for keyword in heading_keywords)

def extract_tables(tables):
    table_count_obj = {"table_count": len(tables), "tables": []}
    for i, table in enumerate(tables):
        table_obj = {"table_number": i + 1, "table_content": [], "table_count": 0}
        rows = table.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        for row in rows:
            cells = row.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            for cell in cells:
                paragraphs = cell.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                for paragraph in paragraphs:
                    if paragraph.text:
                        table_obj["table_content"].append(paragraph.text)
        table_count_obj["tables"].append(table_obj)
    return table_count_obj

def extract_heading_info(file_path, heading_keywords):
    doc = Document(file_path)
    heading_stats = {}
    heading_count = 0

    for paragraph in doc.paragraphs:
        paragraph_text = clean_text(paragraph.text)
        if is_heading(paragraph_text, heading_keywords):
            first_run = paragraph.runs[0] if paragraph.runs else None
            font = first_run.font if first_run else None
            formatting_info = {
                "size": font.size.pt if font and font.size else None,
                "font": font.name if font else None,
                "bold": first_run.bold if first_run else None,
                "italic": first_run.italic if first_run else None,
            }
            if paragraph_text not in heading_stats:
                heading_stats[paragraph_text] = []
            heading_stats[paragraph_text].append(formatting_info)
            heading_count += 1
    
    tables = doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')
    heading_stats["total_headings"] = heading_count
    heading_stats["tables_in_document"] = extract_tables(tables)
    return heading_stats

def run_stats(base_path: Path, manifest_path: str, output_path: str, heading_keywords: list):
    """Main entry point for stats."""
    full_manifest_path = base_path / manifest_path
    all_stats = {}

    if not full_manifest_path.exists():
        print(f"Error: Manifest {full_manifest_path} not found.")
        return False

    with open(full_manifest_path, 'r') as file_list:
        for line in file_list:
            rel_path = line.strip()
            if not rel_path: continue
            file_path = base_path / rel_path
            if file_path.exists():
                all_stats[file_path.name] = extract_heading_info(file_path, heading_keywords)
            else:
                print(f"Warning: File '{file_path}' not found.")

    with open(base_path / output_path, 'w') as f:
        json.dump(all_stats, f, indent=4)
    
    print(f"Document statistics saved to {output_path}")
    return True

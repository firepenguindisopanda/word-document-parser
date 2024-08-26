import os
import json
import unicodedata
from docx import Document

def clean_text(text):
    normalized_text = unicodedata.normalize('NFC', text)
    cleaned_text = ' '.join(normalized_text.split())
    return cleaned_text

def is_heading(text):
    heading_keywords = [
        "Explore Majors",
        "Sample Career",
        "Prospective",
        "Skills And Characteristics",
        "Alternative Career",
        "Employers",
        "Areas",
        "Places that hire",
        "What can I do",
        "Skills Related",
        "Sample Jobs",
        "Sample"
    ]
    return any(text.lower().startswith(keyword.lower()) for keyword in heading_keywords)

def extract_tables(tables):
    table_count_obj = {
        "table_count": len(tables),
        "tables": []
    }
    # iterate through the tables
    for i, table in enumerate(tables):
        # create an object to store the information of the table
        table_obj = {
            "table_number": i + 1,
            "table_content": [],
            "table_count": 0
        }
        # get the rows of the table
        rows = table.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        # iterate through the rows
        for row in rows:
            # get the cells of the row
            cells = row.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            # iterate through the cells
            for cell in cells:
                # get the paragraphs of the cell
                paragraphs = cell.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
                # iterate through the paragraphs
                for paragraph in paragraphs:
                    # get the text of the paragraph
                    text = paragraph.text
                    # check if the text is not empty
                    if text:
                        # append the text to the table content
                        table_obj["table_content"].append(text)
        # append the table object to the tables array
        table_count_obj["tables"].append(table_obj)
    table_count_obj["table_count"] = len(tables)
    # return the table count object
    return table_count_obj

def extract_heading_info(file_path):
    doc = Document(file_path)
    heading_stats = {}
    heading_count = 0

    for paragraph in doc.paragraphs:
        paragraph_text = clean_text(paragraph.text)
        
        if is_heading(paragraph_text):
            print(paragraph_text)  # Print the full heading

            # Get formatting of the first run (assuming consistent formatting)
            first_run = paragraph.runs[0] if paragraph.runs else None
            font = first_run.font if first_run else None

            formatting_info = {
                "size": font.size.pt if font and font.size else None,
                "font": font.name if font else None,
                "bold": first_run.bold if first_run else None,
                "italic": first_run.italic if first_run else None,
                "underline": first_run.underline if first_run else None,
                "color": font.color.rgb if font and font.color and font.color.type == "rgb" else None
            }
            
            if paragraph_text not in heading_stats:
                heading_stats[paragraph_text] = []

            heading_stats[paragraph_text].append(formatting_info)
            heading_count += 1
    tables = doc.element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')
    # print(f'Number of tables: {len(tables)}')

    heading_stats["total_headings"] = heading_count
    heading_stats["tables_in_document"] = extract_tables(tables)
    return heading_stats

def process_documents(file_list_path):
    all_stats = {}

    with open(file_list_path, 'r') as file_list:
        for line in file_list:
            file_path = line.strip()
            if os.path.exists(file_path):
                file_name = os.path.basename(file_path)
                heading_stats = extract_heading_info(file_path)
                all_stats[file_name] = heading_stats
            else:
                print(f"Warning: File '{file_path}' not found.")

    return all_stats

def save_to_json(data, json_file_path):
    with open(json_file_path, 'w') as json_file:
        json.dump(data, json_file, indent=4)

if __name__ == "__main__":
    file_list_path = "valid_docx_paths.txt"
    output_json_path = "headings_stats_v4.json"

    document_stats = process_documents(file_list_path)
    save_to_json(document_stats, output_json_path)

    print(f"Document statistics saved to {output_json_path}")

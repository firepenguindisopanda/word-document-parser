from docxcompose.composer import Composer
from docx import Document
import sys
import os

def convert_path(path):
    if path.startswith('/'):
        path = path[1:]
    return os.path.normpath(path.replace('/', '\\'))


def merge_docs_with_page_breaks(output_path, *input_paths):
    if not input_paths:
        print("No input files provided.")
        return
    base_doc = Document()
    composer = Composer(base_doc)
    for path in input_paths:
        doc = Document(path)
        doc.add_page_break()
        composer.append(doc)
    composer.save(output_path)
    print(f'Merged documents saved to {output_path}')

if __name__ == '__main__':
    output_path = 'merged_what_can_I_do_with_majors_documents.docx'
    with open('valid_docx_paths.txt', 'r') as f:
        input_paths = [convert_path(line.strip()) for line in f if line.strip()]

    if not input_paths:
        print("No valid .docx files found.")
    else:
        merge_docs_with_page_breaks(output_path, *input_paths)

# cite: https://www.freecodecamp.org/news/merge-word-documents-in-python/
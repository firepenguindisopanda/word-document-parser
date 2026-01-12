from docxcompose.composer import Composer
from docx import Document
import os
from pathlib import Path

def run_merger(base_path: Path, manifest_path: str, output_path: str):
    """Main entry point for merger."""
    full_manifest_path = base_path / manifest_path
    
    if not full_manifest_path.exists():
        print(f"Error: Manifest {full_manifest_path} not found.")
        return False
        
    with open(full_manifest_path, 'r') as f:
        input_paths = [base_path / line.strip() for line in f if line.strip()]

    if not input_paths:
        print("No documents found to merge.")
        return False

    # Use the first document as the base
    base_doc = Document(input_paths[0])
    composer = Composer(base_doc)
    
    for path in input_paths[1:]:
        if path.exists():
            doc = Document(path)
            base_doc.add_page_break()
            composer.append(doc)
        else:
            print(f"Warning: File {path} not found. Skipping.")
            
    composer.save(base_path / output_path)
    print(f"Merged documents saved to {output_path}")
    return True

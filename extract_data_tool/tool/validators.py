from pathlib import Path
from rich.console import Console

console = Console()

def check_required_files():
    """Ensure required script files exist in the directory."""
    required_files = ["check_doc_extensions.sh", "extract_data_from_word_docs.py", "stats_of_data_in_folders.py"]
    missing_files = [file for file in required_files if not Path(file).exists()]
    
    if missing_files:
        console.print("[red]The following required files are missing:[/red]")
        for file in missing_files:
            console.print(f"- {file}")
        exit(1)

def check_required_folders():
    """Ensure required folders exist in the directory."""
    required_folders = ["FENG", "FFA", "FHE", "FMS", "FSS_LAW_&_SPORT", "FST"]
    missing_folders = [folder for folder in required_folders if not Path(folder).exists()]
    
    if missing_folders:
        console.print("[red]The following required folders are missing:[/red]")
        for folder in missing_folders:
            console.print(f"- {folder}")
        exit(1)

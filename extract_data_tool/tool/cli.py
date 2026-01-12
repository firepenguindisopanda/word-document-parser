import argparse
from pathlib import Path
from rich.console import Console
from .utils import rename_folders, load_config
from .validators import check_required_folders
from .executor import execute_scripts, execute_merge

console = Console()

def main():
    parser = argparse.ArgumentParser(
        description="Automated Word Document Parser Tool",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument("--path", type=str, default=".", help="Base path to the project directory.")
    parser.add_argument("--rename", action="store_true", help="Rename folders with spaces and commas to underscores.")
    parser.add_argument("--check-folders", action="store_true", help="Check if required folders are present.")
    parser.add_argument("--execute", action="store_true", help="Run full extraction and stats process.")
    parser.add_argument("--merge", action="store_true", help="Merge all documents into one.")
    
    args = parser.parse_args()
    base_path = Path(args.path).resolve()
    
    # Load configuration
    config = load_config(base_path / "config.yaml")
    required_folders = config.get("folders", [])
    heading_keywords = config.get("heading_keywords", [])
    
    if args.rename:
        rename_folders(base_path)
    
    if args.check_folders:
        check_required_folders(base_path, required_folders)
    
    if args.execute:
        check_required_folders(base_path, required_folders)
        execute_scripts(base_path, required_folders, heading_keywords)

    if args.merge:
        execute_merge(base_path)
    
    if not any([args.rename, args.check_folders, args.execute, args.merge]):
        console.print("[bold red]No flags provided. Use --help to see available options.[/bold red]")
        exit(1)

if __name__ == "__main__":
    main()

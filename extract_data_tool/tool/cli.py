import argparse
from rich.console import Console
from tool.utils import rename_folders
from tool.validators import check_required_files, check_required_folders
from tool.executor import execute_scripts

console = Console()

def main():
    parser = argparse.ArgumentParser(
        description="Automates folder renaming, validation, and script execution.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    
    parser.add_argument("--rename", action="store_true", help="Rename folders with spaces and commas to underscores.")
    parser.add_argument("--check-files", action="store_true", help="Check if required script files are present.")
    parser.add_argument("--check-folders", action="store_true", help="Check if required folders are present.")
    parser.add_argument("--execute", action="store_true", help="Run all scripts in sequence.")
    
    args = parser.parse_args()
    
    if args.rename:
        rename_folders()
    
    if args.check_files:
        check_required_files()
    
    if args.check_folders:
        check_required_folders()
    
    if args.execute:
        execute_scripts()
    
    if not any(vars(args).values()):
        console.print("[bold red]No flags provided. Use --help to see available options.[/bold red]")
        exit(1)

if __name__ == "__main__":
    main()

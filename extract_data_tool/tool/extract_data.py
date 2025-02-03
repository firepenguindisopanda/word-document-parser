import os
import argparse
import subprocess
from pathlib import Path
from rich.console import Console
from rich.prompt import Confirm

console = Console()

def command_exists(command: str) -> bool:
    """Check if a command exists on the system."""
    return subprocess.call(["which", command], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL) == 0

def rename_folders():
    """Rename folders with spaces and commas in their names to underscores."""
    console.print("[bold yellow]Checking for folders with spaces or commas in the name...[/bold yellow]")
    
    for folder in Path.cwd().iterdir():
        if folder.is_dir():
            new_name = folder.name.replace(" ", "_").replace(",", "")
            if new_name != folder.name:
                folder.rename(new_name)
                console.print(f"Renamed '{folder.name}' to '{new_name}'")

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

def execute_scripts():
    """Run required scripts in sequence."""
    console.print("[bold green]Running check_doc_extensions.sh...[/bold green]")
    subprocess.run(["bash", "./check_doc_extensions.sh"], check=True)
    
    console.print("[bold green]Running extract_data_from_word_docs.py...[/bold green]")
    subprocess.run(["python3", "./extract_data_from_word_docs.py"], check=True)
    
    console.print("[bold green]Running stats_of_data_in_folders.py...[/bold green]")
    subprocess.run(["python3", "./stats_of_data_in_folders.py"], check=True)

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

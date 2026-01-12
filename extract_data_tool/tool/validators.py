from pathlib import Path
from rich.console import Console

console = Console()

def check_required_folders(base_path: Path, required_folders: list):
    """Ensure required folders exist in the directory."""
    missing_folders = [folder for folder in required_folders if not (base_path / folder).exists()]
    
    if missing_folders:
        console.print("[red]The following required folders are missing:[/red]")
        for folder in missing_folders:
            console.print(f"- {folder}")
        exit(1)
    console.print("[green]All required folders are present.[/green]")

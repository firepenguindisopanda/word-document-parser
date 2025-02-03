import os
import subprocess
from pathlib import Path
from rich.console import Console

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

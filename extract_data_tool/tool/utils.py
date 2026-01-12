import os
import subprocess
import yaml
from pathlib import Path
from rich.console import Console

console = Console()

def load_config(config_path: Path = Path("config.yaml")):
    """Load configuration from a YAML file."""
    if not config_path.exists():
        console.print(f"[bold red]Configuration file not found: {config_path}[/bold red]")
        return {}
    with open(config_path, "r") as f:
        return yaml.safe_load(f)

def rename_folders(base_path: Path = Path.cwd()):
    """Rename folders with spaces and commas in their names to underscores."""
    console.print(f"[bold yellow]Checking for folders with spaces or commas in {base_path}...[/bold yellow]")
    
    for folder in base_path.iterdir():
        if folder.is_dir():
            new_name = folder.name.replace(" ", "_").replace(",", "")
            if new_name != folder.name:
                folder.rename(folder.parent / new_name)
                console.print(f"Renamed '{folder.name}' to '{new_name}'")

def generate_valid_paths(base_path: Path, folders: list, output_file: str = "valid_docx_paths.txt"):
    """
    Search for .docx files in specified folders and generate a manifest file.
    Replaces check_doc_extensions.sh logic.
    """
    console.print("[bold blue]Generating valid document paths...[/bold blue]")
    valid_paths = []
    
    for folder_name in folders:
        folder_path = base_path / folder_name
        if not folder_path.exists():
            console.print(f"  [yellow]Warning: Folder {folder_name} does not exist. Skipping.[/yellow]")
            continue
            
        console.print(f"  Checking folder: {folder_name}")
        for file in folder_path.iterdir():
            if file.is_file() and file.suffix.lower() == ".docx":
                # Rename file if it has spaces
                if " " in file.name:
                    new_filename = file.name.replace(" ", "_")
                    new_file_path = file.parent / new_filename
                    file.rename(new_file_path)
                    file = new_file_path
                    console.print(f"    Renamed: {file.name}")
                
                valid_paths.append(str(file.relative_to(base_path)))
                console.print(f"    Valid: {file.name}")

    with open(base_path / output_file, "w") as f:
        for path in valid_paths:
            f.write(f"{path}\n")
            
    console.print(f"[bold green]Total valid files found: {len(valid_paths)}[/bold green]")
    console.print(f"Manifest saved to {output_file}")
    return valid_paths

import sys
from pathlib import Path
from rich.console import Console
from .utils import generate_valid_paths
from ..core import run_parser, run_stats, run_merger

console = Console()

def execute_scripts(base_path: Path, folders: list, heading_keywords: list):
    """Run required extraction scripts in sequence via direct function calls."""
    
    # 1. Generate the manifest
    generate_valid_paths(base_path, folders)
    
    # 2. Run the extraction logic
    console.print(f"[bold green]Extracting data from documents in {base_path}...[/bold green]")
    parser_success = run_parser(
        base_path=base_path,
        manifest_path="valid_docx_paths.txt",
        output_path="extracted_careers_data.json",
        heading_keywords=heading_keywords
    )
    if not parser_success:
        console.print("[bold red]Extraction failed.[/bold red]")
        sys.exit(1)
    
    # 3. Run the stats logic
    console.print(f"[bold green]Calculating statistics in {base_path}...[/bold green]")
    stats_success = run_stats(
        base_path=base_path,
        manifest_path="valid_docx_paths.txt",
        output_path="headings_stats_v4.json",
        heading_keywords=heading_keywords
    )
    if not stats_success:
        console.print("[bold red]Stats calculation failed.[/bold red]")
        sys.exit(1)

    console.print("[bold green]All processing tasks completed successfully.[/bold green]")

def execute_merge(base_path: Path):
    """Run the document merger."""
    console.print(f"[bold blue]Merging documents in {base_path}...[/bold blue]")
    success = run_merger(
        base_path=base_path,
        manifest_path="valid_docx_paths.txt",
        output_path="merged_what_can_I_do_with_majors_documents.docx"
    )
    if not success:
        console.print("[bold red]Merging failed.[/bold red]")
        sys.exit(1)

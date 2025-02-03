import subprocess
from rich.console import Console

console = Console()

def execute_scripts():
    """Run required scripts in sequence."""
    console.print("[bold green]Running check_doc_extensions.sh...[/bold green]")
    subprocess.run(["bash", "./check_doc_extensions.sh"], check=True)
    
    console.print("[bold green]Running extract_data_from_word_docs.py...[/bold green]")
    subprocess.run(["python3", "./extract_data_from_word_docs.py"], check=True)
    
    console.print("[bold green]Running stats_of_data_in_folders.py...[/bold green]")
    subprocess.run(["python3", "./stats_of_data_in_folders.py"], check=True)

#!/bin/bash

# These folders should be present in the same directory as the script
folders=("FENG" "FFA" "FHE" "FMS" "FSS_LAW_&_SPORT" "FST")

# Set error handling
set -e

# Function to handle errors
handle_error() {
    echo "Error occurred in $1 at line $2"
    exit 1
}

# error handling
trap 'handle_error "$BASH_SOURCE" "$LINENO"' ERR

# Function to check if a command exists
command_exists() {
    command -v "$1" >/dev/null 2>&1
}

# Warning to user that this script will rename folders and files with spaces, so be certain that this is what you want to do
echo -e "\033[1m\033[31mThis script will rename folders and files with spaces in their names to use underscores instead.\033[0m"
echo -e "\033[1m\033[31mPlease ensure that this is what you want to do before proceeding.\033[0m"
echo -e "\033[1m\033[31mYou must close any open files that are related to these folders -> ${folders[@]}.\033[0m"
echo -e "\033[1m\033[31mBecause these folders are expected to be in the same folder as this program.\033[0m"
echo -e "\033[1m\033[31mPress Ctrl+C to cancel or any other key to continue.\033[0m"
read -r


# Check if Bash is installed
if ! command_exists bash; then
    echo "Bash is not installed. Please install Bash and try again."
    exit 1
fi

# Check if Python is installed
if ! command_exists python3; then
    echo "Python 3 is not installed. Please install Python 3 and try again."
    exit 1
fi

# Check if all required files exist
required_files=("check_doc_extensions.sh" "extract_data_from_word_docs.py" "stats_of_data_in_folders.py")
missing_files=()

for file in "${required_files[@]}"; do
    if [ ! -f "$file" ]; then
        missing_files+=("$file")
    fi
done

if [ ${#missing_files[@]} -ne 0 ]; then
    echo "The following required files are missing:"
    for file in "${missing_files[@]}"; do
        echo "- $file"
    done
    echo "Please ensure all required files are present and try again."
    exit 1
fi

# Rename folders with spaces to underscores
# only if there are spaces in the folder names
# check if any folder inside the current directory have spaces then rename them if they do
echo "Checking for folders with spaces in the name, replacing spaces with underscores, and removing any commas..."

# Loop through all directories in the current directory
for folder in */; do
  # Remove the trailing slash from the folder name
  folder_name="${folder%/}"

  # Check if the folder name contains spaces or commas
  if [[ "$folder_name" =~ [[:space:]] || "$folder_name" =~ [,] ]]; then
    # Create a new folder name by replacing spaces with underscores and removing commas
    new_folder_name=$(echo "$folder_name" | tr ' ' '_' | tr -d ',')
    mv "$folder_name" "$new_folder_name"

    echo "Renamed '$folder_name' to '$new_folder_name'"
  fi
done

echo "Folder renaming process completed."


# check if the folders are present in the same directory as the script
echo "Checking if all required folders are present..."
echo "Required folders: ${folders[@]}, Pay attention to the naming of the folders if folders aren't found."
for folder in "${folders[@]}"; do
    if [ ! -d "$folder" ]; then
        echo "Folder $folder not found. Please ensure all required folders are present and try again."
        exit 1
    fi
done

echo "Running check_doc_extensions.sh..."
if ! bash ./check_doc_extensions.sh; then
    echo "check_doc_extension.sh failed. Cannot continue."
    exit 1
fi

echo "Running extract_data_from_word_docs.py..."
python ./extract_data_from_word_docs.py

echo "Running statistics_of_data_in_folders.py..."
python ./stats_of_data_in_folders.py

echo "All tasks completed successfully."
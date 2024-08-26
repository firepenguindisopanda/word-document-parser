#!/bin/bash

output_file="valid_docx_paths.txt"


> "$output_file"


folders=("FENG" "FFA" "FHE" "FMS" "FSS_LAW_&_SPORT" "FST")

count_of_files=0

rename_file() {
    local old_name="$1"
    local dir_name=$(dirname "$old_name")
    local base_name=$(basename "$old_name")
    local name_part="${base_name%.*}"
    local extension="${base_name##*.}"
    
    local new_name_part=$(echo "$name_part" | tr ' ' '_')
    
    local new_name="$dir_name/${new_name_part}.${extension}"
    
    if [ "$old_name" != "$new_name" ]; then
        mv "$old_name" "$new_name"
        echo "Renamed: $old_name -> $new_name" >&2
    fi
    echo "$new_name"
}


for folder in "${folders[@]}"; do
    echo "Checking folder: $folder"
    
    if [ ! -d "$folder" ]; then
        echo "  Warning: Folder $folder does not exist. Skipping."
        continue
    fi
    
    for file in "$folder"/*; do
        if [ -f "$file" ]; then
            if [[ $file == *.docx ]]; then
                renamed_file=$(rename_file "$file")
                echo "  Valid: $renamed_file"
                ((count_of_files++))
                echo "${renamed_file#./}" >> "$output_file"
            else
                echo "  Invalid: $file is not a .docx file"
            fi
        fi
    done
    
    echo "Finished checking $folder"
    echo
done

echo "Total number of valid .docx files: $count_of_files files found"
echo "All folders checked. Valid .docx paths saved to $output_file"
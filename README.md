# word-document-parser
This is a system designed to extract the career data from a set of files in a specific group of folders

[Google Drive Containing the video: How to use Word Document Parser](https://drive.google.com/drive/folders/19KgD7RcTyoxqcowNCyk70GsX_SonMtjU?usp=sharing)

## Running the program

**Make sure you have the following folders in the same directory as the program files**

*The progam will look for these folder*:
("FENG" "FFA" "FHE" "FMS" "FSS LAW & SPORT" "FST")

- The program will rename and folder or files with spaces. It will replace spaces with _ 

The program will generate files as follows:

- extracted_careers_data.json - The data extracted from all 77 word document files
- valid_docx_paths.txt - a list of the valid paths for the parser to find the word documents
- headings_stats_v4.json - a brief insight to the styling of the headings in the word document files

## Steps to run the program

- `pip install -r requirements.txt`
- Ensure the `check_doc_extensions.sh` is executable by using this command -> `chmod +x check_doc_extensions.sh`
- Ensure the `extract_career_data.sh` is executable by using this command -> `chmod +x extract_career_data.sh`
- Make sure the folders are available and that the word document files are not open
- Execute the shell script with this command `./extract_career_data.sh`

# Extracting Career Data From 77 Word Documents

## Initial Investigation

To find all the unique headings, the headings aren't indicated by the format heading in Word, but instead, they are identified by being **bolded**, size 14, and Calibri font.

### Keywords associated with headings in the Word documents:

- "Explore Majors"
- "Sample Career"
- "Prospective"
- "Skills And Characteristics"
- "Alternative Career"
- "Employers"
- "Areas"
- "Places that hire"
- "What can I do"
- "Skills Related"
- "Sample Jobs"
- "Sample"

Each document is separated into a couple of sections as listed above. Below is a breakdown of a few headings:

- **Explore Majors** - What Can I Do With A {degree title}
- **Sample Career Title** For {degree title}
- **Prospective Industries** For {degree title}
- **Skills and Characteristics** Related To A {degree title}

## Working Solution

### Python Library Used

Using the abstracted objects and functions from `python-docx` to get the data from the files was proving difficult. Decided to use the XML format of the files to identify the headings and extract the data from the sections.

### Files that are part of the system:

- **check_doc_extensions.sh**
  - This script checks if there are legitimate Word documents in the required folders.
  - After completion, it generates a file called `valid_docx_paths.txt`.
  - `valid_docx_paths.txt` lists out the paths to the legitimate Word documents in the folders. This will be used in the other Python scripts.

- **extract_data_from_word_docs.py**
  - This script uses the `valid_docx_paths.txt` file to read the Word documents and extract the data.
  - It generates a JSON file named `extracted_careers_data.json` with the extracted data in the format of:
    ```json
    {
        "Word document file.docx": {
            "heading_from_word_doc_file": ["array of string data representing the data for each section"]
        }
    }
    ```

- **stats_of_data_in_folders.py**
  - This script uses the `valid_docx_paths.txt` file to read the Word documents and extract the data.
  - It generates `headings_stats_v4.json` with some data statistics of the files parsed.

- **merge_docs.py**
  - This script uses the `valid_docx_paths.txt` file to read the Word documents and extract the data.
  - It generates a master document of all the Word documents merged, allowing for easier navigation of the Word files.

### Document Formatting

The Word documents are formatted as follows:

- **Section Heading**
  ```json
  {
      "Content"
  }

----

### However, the section headings aren't formatted using the Microsoft Word Heading property. The only way to know that they are headings is by their styling—size 14 bold Calibri.

**Right now, the program is able to extract the heading and the content.**

### **Recommendation**

*Format the headings of sections to have a unique identifying factor, such as using the Heading1, Heading2, or Heading3 properties. This would allow the program to be generalized, only requiring the identification of the heading property.*

**The program is dependent on the headings of sections following this wording schema:**

- "Explore Majors"
- "Sample Career"
- "Prospective"
- "Skills And Characteristics"
- "Alternative Career"
- "Employers"
- "Areas"
- "Places that hire"
- "What can I do"
- "Skills Related"
- "Sample Jobs"
- "Sample"

Any deviation from this or if there is content that follows the same pattern, there will be issues in parsing. Having a uniform format for headings of sections would allow the program to identify any new headings easier and allow the content to be formatted however the Word Document owner decides.

### Unicode Issues Found in the Extracted Text

- \u2013
- \u2019
- \u00a0
- \u201c
- \u201d
- \u00e9
- \u00ee
- \u00f4
- \u2022
- \u2014
- \u2026

### Temporary Fix:

Converted the text to normal English text, e.g., "sauté" is turned to "saute".

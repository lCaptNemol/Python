import os
import glob
import re
from docx import Document
from docxcompose.composer import Composer

def get_common_prefix(filenames):
    """Extracts the prefix before the first '.' from all filenames and ensures they are the same."""
    prefixes = {re.split(r"\.", os.path.basename(f), 1)[0] for f in filenames}
    
    if len(prefixes) == 1:
        return prefixes.pop()
    return "combined_document"  # Default if no common prefix is found

def combine_word_documents(input_folder, output_folder):
    # Ensure output directory exists
    os.makedirs(output_folder, exist_ok=True)
    
    # Get all .docx files while ignoring temp files (~$ files)
    word_files = [
        f for f in glob.glob(os.path.join(input_folder, "*.docx"))
        if not os.path.basename(f).startswith("~$")
    ]

    if not word_files:
        print("No valid Word documents found in the folder.")
        return
    
    # Sort files to maintain order
    word_files.sort()
    
    # Determine output filename dynamically
    output_filename = get_common_prefix(word_files) + ".docx"

    # Open the first document as the base
    master_doc = Document(word_files[0])
    composer = Composer(master_doc)
    
    # Append all other documents
    for file in word_files[1:]:
        print(f"Adding: {file}")  # Debugging line to see which files are processed
        doc = Document(file)
        
        # Insert a page break before adding the next document
        master_doc.add_page_break()
        
        # Append document after page break
        composer.append(doc)
    
    # Define output file path
    output_file_path = os.path.join(output_folder, output_filename)
    
    # Save the final combined document
    composer.save(output_file_path)
    print(f"Merged document saved as: {output_file_path}")
    print(f"Total number of Word documents combined: {len(word_files)}")

if __name__ == "__main__":
    input_folder = "/Users/km/Documents/Combine_Word_Docs/PortHealth"
    output_folder = "/Users/km/Documents/Combine_Word_Docs/Convert"
    combine_word_documents(input_folder, output_folder)
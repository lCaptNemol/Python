import os
import glob
from docx import Document
from docxcompose.composer import Composer

def combine_word_documents(input_folder, output_folder, output_filename="combined_document.docx"):
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
    input_folder = "/Users/km/Documents/Combine_Word_Docs/Importation"
    output_folder = "/Users/km/Documents/Combine_Word_Docs/Convert"
    combine_word_documents(input_folder, output_folder)
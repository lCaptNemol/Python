import os
import glob
import re
from docx import Document
from docxcompose.composer import Composer

BASE_DIR = "/Users/km/Documents/Projects/Combine_Word_Docs"

def list_folders(base_path):
    """Returns a list of subdirectories inside the base path."""
    return [f for f in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, f))]

def get_user_selected_folder(base_path):
    """Prompts the user to select a folder from the available options."""
    folders = list_folders(base_path)
    if not folders:
        print("No subdirectories found in the base directory.")
        return None
    
    print("\nAvailable folders:")
    for i, folder in enumerate(folders, 1):
        print(f"{i}. {folder}")

    while True:
        try:
            choice = int(input("\nEnter the number of the folder to use: "))
            if 1 <= choice <= len(folders):
                return folders[choice - 1]
            else:
                print("Invalid selection. Please choose a valid folder number.")
        except ValueError:
            print("Invalid input. Please enter a number.")

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
    selected_folder = get_user_selected_folder(BASE_DIR)
    
    if selected_folder:
        input_folder = os.path.join(BASE_DIR, selected_folder)
        output_folder = os.path.join(BASE_DIR, "Convert")  # Keep the output in 'Convert' folder
        combine_word_documents(input_folder, output_folder)
    else:
        print("No folder selected. Exiting program.")
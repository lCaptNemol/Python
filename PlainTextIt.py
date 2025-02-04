import os
import json
import docx

def extract_text_from_docx(file_path):
    """Extract text from a Word document (.docx)."""
    doc = docx.Document(file_path)
    text = []
    for para in doc.paragraphs:
        text.append(para.text.strip())
    return "\n".join([t for t in text if t])

def convert_to_json(doc_text, file_name):
    """Convert extracted text into a structured JSON format."""
    json_data = {
        "document_title": file_name,
        "sections": []
    }
    
    sections = doc_text.split("\n\n")  # Split sections based on double newlines
    for section in sections:
        lines = section.split("\n")
        title = lines[0] if lines else "Untitled Section"
        content = " ".join(lines[1:]) if len(lines) > 1 else ""
        json_data["sections"].append({"title": title, "content": content})
    
    return json_data

def process_directory(input_dir, output_dir):
    """Process all .docx files in a directory and convert them to JSON."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    for file in os.listdir(input_dir):
        if file.endswith(".docx"):
            file_path = os.path.join(input_dir, file)
            doc_text = extract_text_from_docx(file_path)
            json_data = convert_to_json(doc_text, file)
            
            json_file_path = os.path.join(output_dir, f"{os.path.splitext(file)[0]}.json")
            with open(json_file_path, "w", encoding="utf-8") as json_file:
                json.dump(json_data, json_file, indent=2)
            
            print(f"Processed: {file} -> {json_file_path}")

# Example usage
input_directory = "/Users/km/Documents/Convert"
output_directory = "/Users/km/Documents/JSON"
process_directory(input_directory, output_directory)

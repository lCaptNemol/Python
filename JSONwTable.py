import docx
import json
import os

def extract_table(table):
    """Converts a Word table into a list of dictionaries (rows)."""
    rows = table.rows
    if not rows:
        return []

    headers = [cell.text.strip() for cell in rows[0].cells]
    table_data = []

    for row in rows[1:]:  # Skip header row
        row_data = {headers[i]: cell.text.strip() for i, cell in enumerate(row.cells)}
        table_data.append(row_data)

    return table_data

def docx_to_json(docx_path, output_path):
    """Converts a Word document (headings, paragraphs, and tables) to JSON."""
    document_data = []

    try:
        doc = docx.Document(docx_path)

        current_section = None
        current_subsection = None
        current_subsubsection = None

        for paragraph in doc.paragraphs:
            style_name = paragraph.style.name
            text = paragraph.text.strip()

            if not text:
                continue

            if style_name.startswith("Heading 1"):
                current_section = {"title": text, "content": []}
                document_data.append(current_section)
                current_subsection = None
                current_subsubsection = None
            elif style_name.startswith("Heading 2"):
                if current_section is None:
                    current_section = {"title": "Untitled Section", "content": []}
                    document_data.append(current_section)

                current_subsection = {"title": text, "content": []}
                current_section["content"].append(current_subsection)
                current_subsubsection = None
            elif style_name.startswith("Heading 3"):
                if current_subsection is None:
                    current_subsection = {"title": "Untitled Subsection", "content": []}
                    current_section["content"].append(current_subsection)

                current_subsubsection = {"title": text, "content": []}
                current_subsection["content"].append(current_subsubsection)
            else:  # Normal or Body Text
                content_item = {"type": "paragraph", "text": text}

                if current_subsubsection:
                    current_subsubsection["content"].append(content_item)
                elif current_subsection:
                    current_subsection["content"].append(content_item)
                elif current_section:
                    current_section["content"].append(content_item)
                else:
                    if not document_data:
                        current_section = {"title": "Introduction", "content": []}
                        document_data.append(current_section)
                    current_section["content"].append(content_item)

        # Process tables
        for table in doc.tables:
            table_data = extract_table(table)

            if table_data:
                table_item = {"type": "table", "data": table_data}

                if current_subsubsection:
                    current_subsubsection["content"].append(table_item)
                elif current_subsection:
                    current_subsection["content"].append(table_item)
                elif current_section:
                    current_section["content"].append(table_item)
                else:
                    if not document_data:
                        current_section = {"title": "Introduction", "content": []}
                        document_data.append(current_section)
                    current_section["content"].append(table_item)

        # Convert to JSON and save
        json_output = json.dumps(document_data, indent=4, ensure_ascii=False)
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(json_output)

        print(f"Conversion successful. JSON output saved to {output_path}")

    except docx.opc.exceptions.PackageNotFoundError:
        print(f"Error: File not found: {docx_path}")
    except Exception as e:
        print(f"An error occurred: {e} while processing {docx_path}")

def process_directory(input_dir, output_dir):
    """Processes all .docx files in the input directory and converts them to JSON."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    for filename in os.listdir(input_dir):
        if filename.endswith(".docx") and not filename.startswith("~$"):  # Ignore temporary Word files
            input_path = os.path.join(input_dir, filename)
            output_filename = os.path.splitext(filename)[0] + ".json"
            output_path = os.path.join(output_dir, output_filename)

            docx_to_json(input_path, output_path)

# ==== SET YOUR INPUT AND OUTPUT DIRECTORY HERE ====
input_directory = "/Users/km/Documents/Convert"  # Change this to your input folder
output_directory = "/Users/km/Documents/JSON"    # Change this to your output folder

process_directory(input_directory, output_directory)
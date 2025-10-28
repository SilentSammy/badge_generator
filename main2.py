from docx import Document
from docx.shared import Inches
import json
import shutil
import os

def replace_text_in_document(doc, replacements):
    """
    Replace text in a Word document like Ctrl+H find-and-replace.
    Preserves formatting by working with runs instead of paragraph.text
    
    Args:
        doc: Document object
        replacements (dict): Dictionary with find:replace pairs
    """
    for find_text, replace_text in replacements.items():
        replaced_count = 0
        
        # Handle multiline replacement text
        if '\n' in replace_text:
            # For multiline text, use Word's line break
            replace_text = replace_text.replace('\n', '\r')
        
        # Replace in all paragraphs
        for paragraph in doc.paragraphs:
            if find_text in paragraph.text:
                replaced_count += replace_in_paragraph(paragraph, find_text, replace_text)
        
        # Replace in all table cells
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if find_text in paragraph.text:
                            replaced_count += replace_in_paragraph(paragraph, find_text, replace_text)
        
        if replaced_count > 0:
            print(f"✅ Replaced '{find_text}' in {replaced_count} location(s)")
        else:
            print(f"⚠️  Warning: '{find_text}' not found in document")

def replace_in_paragraph(paragraph, find_text, replace_text):
    """
    Replace text in a paragraph while preserving formatting by working with runs.
    
    Args:
        paragraph: Word paragraph object
        find_text (str): Text to find
        replace_text (str): Text to replace with
        
    Returns:
        int: Number of replacements made (0 or 1)
    """
    if find_text not in paragraph.text:
        return 0
    
    # Try to replace within individual runs first (preserves formatting)
    for run in paragraph.runs:
        if find_text in run.text:
            run.text = run.text.replace(find_text, replace_text)
            return 1
    
    # If not found in any single run, the text might span multiple runs
    # In this case, we'll use the paragraph.text approach as fallback
    # (This will lose formatting but at least the replacement works)
    if find_text in paragraph.text:
        paragraph.text = paragraph.text.replace(find_text, replace_text)
        return 1
    
    return 0

def replace_images_in_document(doc, image_replacements):
    """
    Replace text placeholders with images in a Word document.
    
    Args:
        doc: Document object
        image_replacements (dict): Dictionary with placeholder:image_path pairs
    """
    for placeholder, image_path in image_replacements.items():
        replaced_count = 0
        
        # Check if image file exists
        if not os.path.exists(image_path):
            print(f"⚠️  Warning: Image file '{image_path}' not found")
            continue
        
        # Replace in all paragraphs
        for paragraph in doc.paragraphs:
            if placeholder in paragraph.text:
                replaced_count += replace_text_with_image_in_paragraph(paragraph, placeholder, image_path)
        
        # Replace in all table cells  
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if placeholder in paragraph.text:
                            replaced_count += replace_text_with_image_in_paragraph(paragraph, placeholder, image_path)
        
        if replaced_count > 0:
            print(f"✅ Replaced '{placeholder}' with image in {replaced_count} location(s)")
        else:
            print(f"⚠️  Warning: '{placeholder}' not found in document")

def replace_text_with_image_in_paragraph(paragraph, placeholder, image_path, width_inches=0.8):
    """
    Replace text placeholder with an image in a paragraph.
    
    Args:
        paragraph: Word paragraph object
        placeholder (str): Text to replace
        image_path (str): Path to replacement image
        width_inches (float): Width of the image in inches
        
    Returns:
        int: Number of replacements made (0 or 1)
    """
    if placeholder not in paragraph.text:
        return 0
    
    # Clear the paragraph and add the image
    paragraph.clear()
    run = paragraph.add_run()
    run.add_picture(image_path, width=Inches(width_inches))
    
    return 1

def table_to_dicts(table, count_per_doc=5, image_column="Image", key_format="{}{:02d}"):
    """
    Split a DataFrame into groups and generate text/image replacement dictionaries.
    
    Args:
        table: pandas DataFrame with attendee data
        count_per_doc (int): Number of attendees per document
        image_column (str): Column name containing image paths
        key_format (str): Format string for generating keys (e.g., "name{:02d}")
        
    Returns:
        list: List of tuples (text_dict, image_dict) for each document group
    """
    result_groups = []
    
    # Split DataFrame into chunks of count_per_doc rows
    for i in range(0, len(table), count_per_doc):
        chunk = table.iloc[i:i + count_per_doc]
        
        text_dict = {}
        image_dict = {}
        
        # Generate dictionaries for this chunk
        for idx, (_, row) in enumerate(chunk.iterrows(), 1):
            # Generate keys for text columns (excluding image column)
            for column in table.columns:
                if column != image_column and column != 'id':  # Skip image and id columns
                    key = key_format.format(column.lower(), idx)
                    text_dict[key] = str(row[column])
            
            # Generate key for image column
            if image_column in table.columns:
                image_key = key_format.format(image_column.lower(), idx)
                image_dict[image_key] = str(row[image_column])
        
        result_groups.append((text_dict, image_dict))
    
    return result_groups

def main():
    """Load JSON data and perform text and image replacements"""
    input_file = "Gafetes.docx"
    output_file = "Gafetes_from_json.docx"
    text_json_file = "dummy_data.json"
    image_json_file = "dummy_images.json"
    
    print("=== Simple Find & Replace from JSON ===")
    print(f"Input: {input_file}")
    print(f"Output: {output_file}")
    print(f"Text data: {text_json_file}")
    print(f"Image data: {image_json_file}")
    
    # Load text replacement JSON data
    try:
        with open(text_json_file, 'r', encoding='utf-8') as f:
            text_replacements = json.load(f)
        print(f"Loaded {len(text_replacements)} text replacements from JSON")
    except Exception as e:
        print(f"❌ Error loading text JSON: {e}")
        return
    
    # Load image replacement JSON data
    try:
        with open(image_json_file, 'r', encoding='utf-8') as f:
            image_replacements = json.load(f)
        print(f"Loaded {len(image_replacements)} image replacements from JSON")
    except Exception as e:
        print(f"❌ Error loading image JSON: {e}")
        return
    
    # Copy and open document
    try:
        shutil.copy2(input_file, output_file)
        doc = Document(output_file)
    except Exception as e:
        print(f"❌ Error opening document: {e}")
        return
    
    # Perform text replacements
    print("\nPerforming text replacements...")
    replace_text_in_document(doc, text_replacements)
    
    # Perform image replacements
    print("\nPerforming image replacements...")
    replace_images_in_document(doc, image_replacements)
    
    # Save document
    try:
        doc.save(output_file)
        print(f"\n🎉 Document saved as: {output_file}")
    except Exception as e:
        print(f"❌ Error saving document: {e}")

if __name__ == "__main__":
    main()
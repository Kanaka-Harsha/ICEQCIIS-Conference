
import zipfile
import xml.etree.ElementTree as ET
import os

def extract_text(docx_path, output_txt_path):
    try:
        with zipfile.ZipFile(docx_path) as docx:
            xml_content = docx.read('word/document.xml')
        
        # Parse XML
        tree = ET.fromstring(xml_content)
        
        # Namespace map
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        text_parts = []
        
        # Iterate over all paragraphs
        for p in tree.findall('.//w:p', ns):
            # Iterate over all text elements within the paragraph
            # We use .//w:t to get all text nodes in the paragraph, 
            # preserving order if possible (though findall on element returns in document order)
            texts = [node.text for node in p.findall('.//w:t', ns) if node.text]
            if texts:
                text_parts.append(''.join(texts))
            else:
                # Add a newline for empty paragraphs to preserve structure
                text_parts.append('') 
                
        full_text = '\n'.join(text_parts)
        
        with open(output_txt_path, 'w', encoding='utf-8') as f:
            f.write(full_text)
            
        return True

    except Exception as e:
        print(f"Error: {e}")
        return False

if __name__ == "__main__":
    import sys
    if len(sys.argv) > 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
        if extract_text(input_file, output_file):
            print("Done")
        else:
            print("Failed")
    else:
        if extract_text("Content.docx", "extracted_content.txt"):
            print("Done")
        else:
            print("Failed")

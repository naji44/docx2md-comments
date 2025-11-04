import sys
import zipfile
import xml.etree.ElementTree as ET
from docx import Document

def extract_comments_with_locations(docx_path):
    comments = {}
    with zipfile.ZipFile(docx_path) as docx_zip:
        if 'word/comments.xml' in docx_zip.namelist():
            comments_xml = docx_zip.read('word/comments.xml')
            root = ET.fromstring(comments_xml)
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            for comment in root.findall('.//w:comment', ns):
                comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                text_parts = []
                for para in comment.findall('.//w:t', ns):
                    if para.text:
                        text_parts.append(para.text)
                comments[comment_id] = ''.join(text_parts)
    
    doc = Document(docx_path)
    result = []
    
    for para in doc.paragraphs:
        para_text = ""
        comment_positions = []
        
        current_pos = 0
        for run in para.runs:
            run_text = run.text
            run_xml_str = run._element.xml.decode('utf-8') if isinstance(run._element.xml, bytes) else str(run._element.xml)
            
            if 'commentReference' in run_xml_str:
                ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                for elem in run._element.iter():
                    if 'commentReference' in str(elem.tag):
                        cid = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                        if cid and cid in comments:
                            comment_positions.append((current_pos + len(run_text), cid))
            
            para_text += run_text
            current_pos += len(run_text)
        
        if comment_positions:
            comment_positions.sort(reverse=True)
            for pos, cid in comment_positions:
                comment_text = comments[cid]
                para_text = para_text[:pos] + f" **[FEEDBACK: {comment_text}]**" + para_text[pos:]
        
        if para_text.strip():
            result.append(para_text)
    
    return result

def convert_to_markdown(docx_path, output_path):
    lines = extract_comments_with_locations(docx_path)
    filename = docx_path.split('\\')[-1].replace('.docx', '')
    md_content = f"# {filename}\n\n" + '\n\n'.join(lines)
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(md_content)

if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(1)
    
    docx_file = sys.argv[1]
    output_file = docx_file.replace('.docx', '_feedback.md')
    convert_to_markdown(docx_file, output_file)

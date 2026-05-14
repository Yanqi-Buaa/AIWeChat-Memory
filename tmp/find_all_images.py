import docx
from lxml import etree

path = r"C:\Users\Administrator\cow\tmp\（公开）附件3：成果应用情况审查报告-20260203.docx"
doc = docx.Document(path)

nsmap = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
}

# Search entire document body for image references
body = doc.element.body
xml_str = etree.tostring(body, pretty_print=True).decode()

# Find all embed attribute values
import re
# Find all r:embed values
embeds = re.findall(r'r:embed="([^"]+)"', xml_str)
print(f"Total embed references found: {len(embeds)}")
for e in embeds:
    if e in doc.part.rels:
        rel = doc.part.rels[e]
        print(f"  {e} -> {rel.target_ref}")
    else:
        print(f"  {e} -> NOT FOUND in rels")

# Also find all image rId references
for rid, rel in doc.part.rels.items():
    if "image" in str(rel.reltype).lower():
        print(f"\n  Known image: {rid} -> {rel.target_ref}")

# Also search shape fills
fills = re.findall(r'fill="([^"]+)"', xml_str)
print(f"\nImage-like fills: {[f for f in fills if 'image' in f.lower()]}")

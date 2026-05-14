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
}

# Find all paragraphs with images
for i, p in enumerate(doc.paragraphs):
    # Check for drawing elements (new format)
    drawings = p._element.findall('.//w:drawing', nsmap)
    if drawings:
        print(f"[{i}] drawing elements: {len(drawings)}")
        for d in drawings:
            # Find image reference
            blips = d.findall('.//a:blip', nsmap)
            for blip in blips:
                embed = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                if embed:
                    rel = doc.part.rels[embed]
                    print(f"      Image: {rel.target_ref} (reltype={rel.reltype})")
    
    # Check for pict elements (old format)
    picts = p._element.findall('.//w:pict', nsmap)
    if picts:
        print(f"[{i}] pict elements: {len(picts)}")
        for pict in picts:
            imagedatas = pict.findall('.//w:imagedata', nsmap)
            for img in imagedatas:
                embed = img.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                if embed:
                    rel = doc.part.rels[embed]
                    print(f"      Image: {rel.target_ref}")
    
    # Also check for inline shapes
    inlines = p._element.findall('.//wp:inline', nsmap)
    if inlines:
        print(f"[{i}] inline shapes: {len(inlines)}")

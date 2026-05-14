import docx
from lxml import etree

path = r"C:\Users\Administrator\cow\tmp\（公开）附件3：成果应用情况审查报告-20260203.docx"
doc = docx.Document(path)

body = doc.element.body
xml_str = etree.tostring(body, pretty_print=True).decode('utf-8')

# Find all rId references
import re
# Search for any attribute containing rId
rids = set()
for match in re.finditer(r'rId\d+', xml_str):
    rids.add(match.group())
    
print(f"All rId references found in body: {len(rids)}")
for rid in sorted(rids):
    if rid in doc.part.rels:
        rel = doc.part.rels[rid]
        print(f"  {rid} -> {rel.target_ref}")
    else:
        print(f"  {rid} -> NOT IN RELS")

# Let's look at the actual XML around each image reference
# Find 'media/' references
for match in re.finditer(r'media/[^"]+', xml_str):
    print(f"\n  Found media ref: {match.group()}")
    # Show context
    start = max(0, match.start() - 200)
    end = min(len(xml_str), match.end() + 200)
    context = xml_str[start:end]
    print(f"  Context: ...{context}...")

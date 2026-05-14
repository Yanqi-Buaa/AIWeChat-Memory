import docx, os

path = r"C:\Users\Administrator\cow\tmp\（公开）附件4：应用证明模板-20260203.docx"
doc = docx.Document(path)

for ti, table in enumerate(doc.tables):
    print(f"\n--- Table {ti} ({len(table.rows)} x {len(table.columns)}) ---")
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            txt = cell.text.strip()
            if txt:
                print(f"\n  Row {ri} Col {ci}:\n{txt}")

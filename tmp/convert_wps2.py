import os, sys
import win32com.client as win32

base = r"A:\0-学习-2025.2-\A 2023.1 三批lj申请\00-飞行状态\结题\C4.4课题验收\应用证明\应用评价证明"
output_dir = r"C:\Users\Administrator\cow\tmp"

files = [
    "（公开）附件1：应用证明需求申请表-20260203.wps",
    "（公开）附件3：成果应用情况审查报告-20260203.wps", 
    "（公开）附件4：应用证明模板-20260203.wps"
]

try:
    word = win32.gencache.EnsureDispatch('Word.Application')
    print("Word.Application started successfully")
    word.Visible = False
    word.DisplayAlerts = False
    
    for fname in files:
        src = os.path.join(base, fname)
        dst = os.path.join(output_dir, fname.rsplit('.', 1)[0] + '.docx')
        print(f"Converting: {fname}")
        doc = word.Documents.Open(src)
        print(f"  Opened: {doc.Name}")
        doc.SaveAs(dst, 16)  # 16 = wdFormatXMLDocument (docx)
        print(f"  Saved: {dst}")
        doc.Close()
    
    word.Quit()
    print("Done!")
except Exception as e:
    print(f"Error: {e}", file=sys.stderr)
    import traceback
    traceback.print_exc()
    try:
        word.Quit()
    except:
        pass

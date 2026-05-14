import os, zipfile, shutil

src_path = r"C:\Users\Administrator\cow\tmp\（公开）附件3：成果应用情况审查报告-20260203.docx"
out_dir = r"C:\Users\Administrator\cow\tmp\images"
os.makedirs(out_dir, exist_ok=True)

# Extract images from docx (which is a ZIP file)
with zipfile.ZipFile(src_path, 'r') as z:
    for name in z.namelist():
        if name.startswith('word/media/'):
            basename = os.path.basename(name)
            ext = os.path.splitext(basename)[1].lower()
            # Determine the order based on the file number
            z.extract(name, out_dir)
            # Also copy to flat dir for easy access
            shutil.copy(os.path.join(out_dir, name), os.path.join(out_dir, basename))
            print(f"  Extracted: {basename} ({os.path.getsize(os.path.join(out_dir, basename))} bytes)")

# List all images in order
print("\nImages in order:")
for f in sorted(os.listdir(out_dir)):
    if os.path.isfile(os.path.join(out_dir, f)) and not f.endswith('.xml') and 'word' not in f:
        print(f"  {f}")

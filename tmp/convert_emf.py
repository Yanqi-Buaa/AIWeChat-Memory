"""Convert EMF files to PNG using win32com (PowerPoint handles EMF well)"""
import os
import win32com.client as win32

img_dir = r"C:\Users\Administrator\cow\tmp\images"

# Use PowerPoint to convert EMF to PNG
pp = win32.gencache.EnsureDispatch('PowerPoint.Application')
pp.Visible = False
pp.DisplayAlerts = False

for f in os.listdir(img_dir):
    if f.lower().endswith('.emf'):
        emf_path = os.path.join(img_dir, f)
        png_path = os.path.join(img_dir, f.replace('.emf', '.png'))
        if not os.path.exists(png_path):
            print(f"Converting: {f}")
            try:
                # Create a temporary presentation with the EMF
                pres = pp.Presentations.Add()
                slide = pres.Slides.Add(1, 1)  # ppLayoutBlank
                pic = slide.Shapes.AddPicture(emf_path, False, True, 0, 0)
                # Export as PNG
                pic.Export(png_path, 'png')
                pres.Close()
                print(f"  -> {os.path.basename(png_path)}")
            except Exception as e:
                print(f"  Error: {e}")

pp.Quit()
print("\nDone!")

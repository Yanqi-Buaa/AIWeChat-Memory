import os, subprocess, glob

# Check common install paths
paths = [
    r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE',
    r'C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE',
    r'C:\Program Files\WPS Office\**\wps.exe',
]
for p in paths:
    if '**' in p:
        found = glob.glob(p)
        if found:
            for f in found:
                print(f'Found: {f}')
        else:
            print(f'Not found: {p}')
    else:
        print(f'Exists={os.path.exists(p)}: {p}')

# Check associated program for .wps
import winreg
try:
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, '.wps') as key:
        val = winreg.QueryValue(key, None)
        print(f'.wps association: {val}')
except Exception as e:
    print(f'No .wps association: {e}')

# Check if python has any text extraction libs
for mod in ['olefile', 'antiword', 'textract', 'pdfplumber', 'win32file']:
    try:
        __import__(mod)
        print(f'{mod} available')
    except ImportError:
        print(f'{mod} not available')

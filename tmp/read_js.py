import sys

try:
    with open('C:/Users/Administrator/cow/tmp/console_copy.js', 'rb') as f:
        raw = f.read()
        lines = raw.split(b'\n')
        print(f'Total lines: {len(lines)}', file=sys.stderr)
        
        # Find renderMarkdown function
        for i, line in enumerate(lines):
            if b'function renderMarkdown' in line:
                print(f'Found renderMarkdown at line {i+1}', file=sys.stderr)
                for j in range(i, min(i+25, len(lines))):
                    text = lines[j].decode('utf-8', errors='replace')
                    print(f'{j+1}: {text}', file=sys.stderr)
                break
except Exception as e:
    print(f'Error: {e}', file=sys.stderr)

sys.stderr.flush()

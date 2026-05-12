import docx

DST = r'A:\0-学习-2025.2-\A 2023.1 三批lj申请\00-飞行状态\结题\C4.4课题验收\20260429\测试大纲\11.4-项目技术指标测试报告-北航-20260512-修改版-已修改-重点标注.docx'
doc = docx.Document(DST)

# Check specific paragraphs for bold/underline
check = [5, 33, 61, 64, 67, 71, 75, 78, 80, 83, 100, 109, 110, 112]

with open(r'C:\Users\Administrator\cow\tmp\verify_hl.txt', 'w', encoding='utf-8') as f:
    for i in check:
        p = doc.paragraphs[i]
        f.write(f'=== P{i} ===\n')
        for j, r in enumerate(p.runs):
            if r.text.strip():
                bold = 'B' if r.bold else '_'
                ul = 'U' if r.underline else '_'
                try:
                    color = str(r.font.color.rgb) if r.font.color and r.font.color.rgb else 'default'
                except:
                    color = 'err'
                f.write(f'  R{j} [{bold}{ul} {color}]: {r.text[:120]}\n')
        f.write('\n')

print('done')

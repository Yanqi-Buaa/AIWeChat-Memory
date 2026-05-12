import docx, sys

DST = r'A:\0-学习-2025.2-\A 2023.1 三批lj申请\00-飞行状态\结题\C4.4课题验收\20260429\测试大纲\11.4-项目技术指标测试报告-北航-20260512-修改版-已修改.docx'
doc = docx.Document(DST)

check_paras = [75, 80, 83, 84, 86, 87, 98, 99, 100, 113, 121, 122, 123, 124]

with open(r'C:\Users\Administrator\cow\tmp\verify.txt', 'w', encoding='utf-8') as f:
    for i in check_paras:
        p = doc.paragraphs[i]
        t = p.text.strip()
        has_red = False
        for r in p.runs:
            if r.font.color and r.font.color.rgb:
                if str(r.font.color.rgb) == 'FF0000':
                    has_red = True
                    break
        f.write(f'P{i} [red={has_red}]: {t[:200]}\n\n')
print('done')

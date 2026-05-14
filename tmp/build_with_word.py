# -*- coding: utf-8 -*-
"""
Use Word COM to:
1. Cut figure captions, cross-references and images from end
2. Paste at correct positions
3. Write body text with proper paragraph breaks
"""
import os, sys, time
import win32com.client as win32

target = r"A:\0-学习-2025.2-\A 2023.1 三批lj申请\00-飞行状态\结题\C4.4课题验收\应用证明\应用评价证明\基础研究成果应用评价证明.docx"

word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False
word.DisplayAlerts = False

try:
    # Use Selection for document manipulation
    wdDoc = word.Documents.Open(target)
    sel = word.Selection
    
    # Constants
    wdGoToLine = 3
    wdGoToAbsolute = 1
    wdCollapseEnd = 0
    wdStory = 6  # wdGoToStory
    wdLine = 5
    
    def go_to_paragraph_text(text_fragment, match_start=True):
        """Move cursor to the paragraph containing text_fragment."""
        sel.HomeKey(Unit=6)  # wdStory
        find = sel.Find
        find.ClearFormatting()
        find.Text = text_fragment
        find.Forward = True
        find.Wrap = 1
        find.MatchCase = False
        find.Execute()
        return sel.Find.Found
    
    def insert_text(text, new_paragraph=True):
        """Insert text at current selection."""
        if new_paragraph:
            sel.TypeParagraph()
        sel.TypeText(text)
    
    # ======= SECTION 1: Write body text =======
    
    # Go to paragraph [30] - 概述 body
    go_to_paragraph_text('介绍成果应用背景及具体对象')
    sel.Select()  # Select the whole line
    sel.TypeText('')  # Clear it
    overview = (
        '本项目紧密围绕航空发动机整机振动突出、耦合因素多、机理复杂等工程难题，'
        '以某高推重比涡扇发动机为应用对象，开展复杂机动飞行状态下航空发动机转子-支承-机匣系统结构动力学特性研究。'
        '项目突破了气动、机动惯性等复杂载荷作用下发动机结构动力学建模和分析技术，'
        '系统掌握了整机耦合振动机理和动载荷传递规律，探明了机匣连接结构、主轴承及支承构件界面等非线性因素对'
        '转子-支承-机匣系统动力学特性的影响机理和规律，建立了相应的整机结构动力学设计理论与方法，'
        '形成了三项核心成果：'
    )
    sel.TypeText(overview)
    insert_text('（1）复杂飞行条件下整机结构系统动力特性仿真与评估技术；')
    insert_text('（2）基于飞-发多源数据融合的整机振动特征提取与状态识别技术；')
    insert_text('（3）不同状态下高推重比发动机及大功率涡桨发动机典型振动特征参数数据库。')
    sel.TypeParagraph()
    sel.TypeText('上述成果已成功应用于某高推重比涡扇发动机的研制、试飞及交付全过程，'
        '为整机变形评估与振动数据分析提供了有力支撑，应用效果良好。')
    
    print("✅ 概述 written")
    
    # Go to [32] empty paragraph under "成果1的应用情况及效果"
    go_to_paragraph_text('成果1的应用情况及效果')
    sel.MoveDown(Unit=wdLine, Count=1)  # Move to next line (the empty body)
    
    cg1 = (
        '该项成果主要针对某型小涵道比涡扇发动机的研制需求，重点围绕其机动飞行环境及复杂的进排气条件，'
        '建立了结构特征等效的整机有限元模型，形成了覆盖不同整机运动状态和工作状态的振动响应仿真流程。'
        '在应用过程中，创新性地考虑了机动惯性载荷和气动载荷对整机结构变形的耦合效应，'
        '提出了支承约束力学特性的等效施加方法，突破了传统仿真未充分考虑变形引致支承特性改变的局限。'
    )
    sel.TypeText(cg1)
    sel.TypeParagraph()
    sel.TypeText('依托该技术，全面开展了五大类核心仿真分析工作：整机结构建模及力学特性分析、'
        '机动惯性/气动载荷下整机变形仿真、机动惯性载荷下整机振动响应仿真、'
        '加力-喷管气动激励下振动响应特征分析，以及气流旋转激励下的振动响应分析。')
    sel.TypeParagraph()
    # Insert cross-reference to 图1 and 图2 here - will come from the end section
    
    print("✅ 成果1 body written")
    
    # Go to [34] under "成果2的应用情况及效果"
    go_to_paragraph_text('成果2的应用情况及效果')
    sel.MoveDown(Unit=wdLine, Count=1)
    
    cg2_1 = (
        '该项成果主要应用于地面台架、高空台及外场飞行试验全过程的监视与诊断。'
        '某高推重比涡扇发动机是我国第一型完全独立自主研制的涡扇发动机，'
        '其工作包线范围、推重比处于世界领先水平。'
        '在外场试飞过程中，曾出现出厂振动合格的发动机左发整机振动幅值异常增大并逼近限值，'
        '而右发正常的现象，'
    )
    sel.TypeText(cg2_1)
    # Cross-ref to 图3 will go here
    
    print("✅ 成果2 body written")
    
    # Go to [36] under "成果3的应用情况及效果"
    go_to_paragraph_text('成果3的应用情况及效果')
    sel.MoveDown(Unit=wdLine, Count=1)
    
    cg3 = (
        '该项成果采用短时傅里叶变换实现高频采样信号的频域转化与数据压缩，'
        '提取转速基频、模态频率及频域能量分布特征。'
        '基于成果中的整机特征提取与分析技术，利用四阶张量对发动机结构状态特征进行多维描述'
        '（架次-切片-参数-统计量）。'
        '通过对多台同型号发动机历史数据的集成，'
        '构建了\u201c不同状态下高推重比发动机及大功率涡桨发动机典型振动特征参数数据库\u201d。'
        '该数据库基于实测数据提炼关键特征参数，直接服务于飞行载荷确定及振动基准建立，'
    )
    sel.TypeText(cg3)
    # Cross-ref to 图6 will go here
    
    print("✅ 成果3 body written")
    
    # Go to [38] under "成果应用结论"
    go_to_paragraph_text('成果应用结论')
    sel.MoveDown(Unit=wdLine, Count=1)
    
    conclusion = (
        '本项目形成的\u201c复杂机动飞行状态下航空发动机整机动力特性评估、结构设计与监视诊断技术\u201d，'
        '从仿真评估、状态监视识别到数据基准构建，'
        '形成了一套自主可控的面向真实服役环境的航空发动机整机结构动力特性评估与状态识别体系。'
    )
    sel.TypeText(conclusion)
    insert_text('三项核心成果均已成功应用于某高推重比涡扇发动机的研制全过程：')
    insert_text('成果一（复杂飞行条件下整机结构系统动力特性仿真与评估技术）'
        '服务于研发设计与验证阶段，支撑了结构薄弱环节识别与整机动力学设计优化；')
    insert_text('成果二（基于飞-发多源数据融合的整机振动特征提取与状态识别技术）'
        '应用于地面台架与外场试飞全过程的监视诊断，有效识别多起潜在异常并指导装配调整；')
    insert_text('成果三（不同状态下高推重比发动机及大功率涡桨发动机典型振动特征参数数据库）'
        '服务于全寿命周期健康管理，建立了整机振动水平基准与量化评估标准。')
    sel.TypeParagraph()
    sel.TypeText('综上所述，本项目成果在型号研制与生产中应用效果良好，'
        '有效提升了复杂服役环境下航空发动机整机结构动力学评估与状态诊断能力，'
        '保障了装备的研制进度与运行安全，具有广阔的工程应用前景。')
    
    print("✅ 结论 written")
    
    # ======= Save =======
    wdDoc.Save()
    print("\n✅ Document saved successfully!")
    
except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
finally:
    word.Quit()

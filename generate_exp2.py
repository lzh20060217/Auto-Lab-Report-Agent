"""实验二 功率放大器 —— 关键字检索生成"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from skill.sysu_lab_skill import *

def build():
    exp_title = '实验二  功率放大器'
    doc = init_document(experiment_title=exp_title)

    add_centered_run_paragraph(doc, '高频电路实验报告', cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_MAIN_TITLE, bold=True, space_before=4, space_after=4)
    add_info_table(doc)
    add_horizontal_rule(doc, space_before=6, space_after=8)
    add_centered_run_paragraph(doc, exp_title, cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_EXP_TITLE, bold=True, space_before=4, space_after=8)

    add_section_title(doc, '一、实验目的')
    for i, obj in enumerate([
        "加深对丙类功率放大器基本工作原理的理解，掌握丙类功率放大器的调谐特性。",
        "掌握输入激励电压、集电极电源电压及负载变化对放大器工作状态的影响。",
        "通过实验了解基极调幅的工作原理与调幅波特征。",
        "学会测量高频功率放大器的输出功率与效率，理解功放在发射系统中的核心地位。",
    ], 1):
        add_left_paragraph(doc, f'{i}. {obj}')

    add_section_title(doc, '二、实验原理')
    add_sub_section_title(doc, '2.1  高频功率放大器概述')
    add_left_paragraph(doc,
        '高频功率放大器将直流能量转换为高频交流输出，是通信发送装置的核心部件。'
        '与小信号放大器不同，其输入信号幅度达几百毫伏至几伏，晶体管工作延伸至'
        '非线性区（截止与饱和区），追求高输出功率与高效率，通常工作于丙类状态。')
    add_sub_section_title(doc, '2.2  导通角与放大器分类')
    add_left_paragraph(doc,
        '通角 θ 为集电极电流流通角度的一半。甲类 θ=180°，效率约 50%；乙类 θ=90°，'
        '效率约 78%；丙类 θ<90°，效率最高。丙类放大器采用反向偏置，静态时截止，'
        '仅激励信号正半周幅度足够大时导通，集电极电流为周期性余弦脉冲。')
    add_sub_section_title(doc, '2.3  谐振负载与选频滤波')
    add_left_paragraph(doc,
        '丙类功放的集电极电流虽为余弦脉冲，但傅里叶展开后含直流、基波和各次谐波。'
        'LC 谐振回路调谐于基波频率，对基波呈大阻抗选出完整正弦波，对直流和谐波'
        '呈小阻抗予以滤除，同时起到阻抗匹配作用。')
    add_sub_section_title(doc, '2.4  三种工作状态')
    add_left_paragraph(doc,
        '欠压：始终在放大区，电流为尖顶余弦脉冲，输出功率较小。临界：刚进入饱和区'
        '边缘，输出功率最大，为最佳工作点。过压：部分时间进入饱和区，电流顶部凹陷'
        '呈马鞍形，输出电压最大但功率略降。状态由 Ec、Eb、Ubm 和 Rc 决定。')
    add_image_placeholder_box(doc, 1, '丙类高频功率放大器原理电路图')
    add_figure_caption(doc, '丙类高频功率放大器原理电路图', 1)
    add_sub_section_title(doc, '2.5  实验电路')
    add_left_paragraph(doc,
        '三级放大：3Q1 甲类前置（电阻负载无选频），3Q2 丙类功放（3K1 切换 6.3MHz/2MHz '
        '谐振回路），3Q3 丙类末级（6.3MHz）。3W3 调集电极电压，3W4 调负载电阻，'
        '3K2 控制负载接入。')
    add_image_placeholder_box(doc, 2, '高频功率放大器实验电路图')
    add_figure_caption(doc, '高频功率放大器实验电路图', 2)

    add_section_title(doc, '三、实验数据记录')
    add_sub_section_title(doc, '3.1  前置放大级测试')
    add_data_table(doc, ['测试点', '输入/mV(p-p)', '输出/mV(p-p)', '放大倍数 Av'],
                   [['3TP2→3TP3', '——', '——', '——']])
    add_left_paragraph(doc, '注：信号源 6.3MHz，约 400mV(p-p)。')

    add_sub_section_title(doc, '3.2  激励电压 Ub 对工作状态的影响')
    add_data_table(doc, ['Ub/mV(p-p)', 'Ec/V', 'RL/kΩ', '3TP5波形', '状态'],
                   [['——','4','——','——','欠压'],['——','4','——','——','临界'],['——','4','——','——','过压']])
    add_image_placeholder_box(doc, 3, '三种状态下 3TP5 电流脉冲波形')
    add_figure_caption(doc, '欠压、临界、过压状态下 3TP5 电流脉冲波形', 3)

    add_sub_section_title(doc, '3.3  集电极电压 Ec 对工作状态的影响')
    add_data_table(doc, ['Ec/V', 'Ub/mV(p-p)', 'RL/kΩ', '3TP5波形', '状态'],
                   [['——','360','4','——','——'] for _ in range(5)])
    add_left_paragraph(doc, '注：Ec 在 3~7V 范围变化。')

    add_sub_section_title(doc, '3.4  负载电阻 RL 对工作状态的影响')
    add_data_table(doc, ['RL/kΩ', 'Ec/V', 'Ub/mV(p-p)', '3TP5波形', '状态'],
                   [['——','4','360','——','欠压'],['——','4','360','——','临界'],['——','4','360','——','过压']])

    add_sub_section_title(doc, '3.5  功放调谐特性测试')
    add_data_table(doc, ['f/MHz','5.3','5.5','5.7','5.9','6.1','6.3','6.5','6.7','6.9','7.1','7.3'],
                   [['Vc/V(p-p)']+['——']*11])
    add_left_paragraph(doc, '注：Ub=400mV(p-p)，以 6.3MHz 为中心，间隔 200kHz。')
    add_image_placeholder_box(doc, 4, '功放 Vc-f 调谐特性曲线')
    add_figure_caption(doc, '功放 Vc-f 调谐特性曲线', 4)

    add_sub_section_title(doc, '3.6  基极调幅波形观察')
    add_data_table(doc, ['项目','记录内容'],
                   [['调制频率/kHz','——'],['调制幅度/mV(p-p)','——'],['调幅波Umax/V(p-p)','——'],
                    ['调幅波Umin/V(p-p)','——'],['调幅度ma','——'],['波形特征','——']])
    add_image_placeholder_box(doc, 5, '基极调幅输出 3TP4 波形')
    add_figure_caption(doc, '基极调幅输出 3TP4 波形', 5)

    add_sub_section_title(doc, '3.7  输出功率测试')
    add_data_table(doc, ['项目','记录值'],
                   [['频率/MHz','6.3'],['3R02振幅/V(p-p)','——'],['有效值U/V','——'],
                    ['P=U²/100/W','——'],['3Q3集电极电流/mA','——'],['直流输入功率/mW','——'],['效率η/%','——']])
    add_left_paragraph(doc, '注：3R02=100Ω。')

    add_section_title(doc, '四、结果分析与总结')
    add_sub_section_title(doc, '4.1  工作状态综合分析')
    add_left_paragraph(doc, '根据 3.2~3.4 数据，分析 Ub、Ec、RL 对三种工作状态的转换规律，'
                       '说明各状态下电流波形和输出功率的变化趋势。')
    add_left_paragraph(doc, '分析：')
    for _ in range(4): add_left_paragraph(doc, '')
    add_sub_section_title(doc, '4.2  调谐特性分析')
    add_left_paragraph(doc, '绘制 Vc-f 曲线，求 f₀ 和 −3dB 带宽，分析 LC 回路的选频滤波作用。')
    add_left_paragraph(doc, '分析：')
    for _ in range(4): add_left_paragraph(doc, '')
    add_sub_section_title(doc, '4.3  效率分析')
    add_left_paragraph(doc, '计算功放效率 η，分析丙类高效率的物理根源——导通角小、直流分量占比低。'
                       '讨论效率与线性度之间的工程折衷。')
    add_left_paragraph(doc, '分析：')
    for _ in range(4): add_left_paragraph(doc, '')

    add_section_title(doc, '五、实验总结')
    add_left_paragraph(doc,
        '本次功率放大器实验让我对丙类功放的物理本质有了深刻认识。丙类放大器区别于甲类'
        '的核心在于导通角 θ<90°，晶体管仅在激励信号正半周且超过反偏后才短暂导通，'
        '集电极电流为周期性余弦脉冲。实验中发现，通过调整激励电压 Ub、集电极电压 Ec '
        '和负载电阻 RL，可以灵活切换欠压、临界和过压三种工作状态——欠压时电流为尖顶'
        '脉冲、输出功率较小；临界时输出功率达到最大，是实际设计的目标工作点；过压时'
        '电流顶部出现凹陷，输出幅度虽大但效率有所折损。这一观察直观验证了理论课中关于'
        '动态负载线的分析。调谐特性测试进一步展示了 LC 谐振回路既选频又匹配的双重功能，'
        '功率测试环节也让我切身体会到丙类功放以牺牲线性度换取高效率的工程哲学——这一'
        '特性使其天然适配恒包络调制信号（如 FM），而处理包络变化的 AM 信号时则需额外'
        '补偿。本次实验成功串联了导通角、傅里叶分析、谐振负载和阻抗匹配等核心概念，'
        '为理解无线电发射机末级功放的工程设计打下了坚实基础。')

    return doc

def main():
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(d, exist_ok=True)
    doc = build()
    dp = os.path.join(d, '高频电路实验_实验2_功率放大器.docx')
    doc.save(dp)
    print(f'DOCX: {dp}')
    from docx2pdf import convert
    try:
        convert(dp, dp.replace('.docx','.pdf'))
        print(f'PDF: {dp.replace(".docx",".pdf")}')
    except Exception as e:
        print(f'PDF失败: {e}')

if __name__ == '__main__':
    main()

"""
实验报告优化版生成器 v3 —— 严格调用 SYSU_Lab_Report_Gen 排版技能

排版铁律见 skill/sysu_lab_skill.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from skill.sysu_lab_skill import (
    SKILL,
    init_document,
    add_info_table,
    add_horizontal_rule,
    add_centered_run_paragraph,
    add_left_paragraph,
    add_section_title,
    add_sub_section_title,
    add_data_table,
    add_image_placeholder_box,
    add_figure_caption,
    add_blank_line,
    setup_page,
    setup_header,
    setup_footer,
    set_run_font,
)


def build_report():
    exp_title = '实验一 小信号调谐放大器（单调谐与双调谐放大器）'
    doc = init_document(experiment_title=exp_title)

    # === 顶部大标题 ===
    add_centered_run_paragraph(doc, '高频电路实验报告',
                               cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_MAIN_TITLE,
                               bold=True,
                               space_before=4, space_after=4)

    # === 基础信息区（3行×2列 无边框隐形表格，绝对对齐）===
    add_info_table(doc)

    # === 水平分割线 ===
    add_horizontal_rule(doc, space_before=6, space_after=8)

    # === 实验小标题 ===
    add_centered_run_paragraph(doc, exp_title,
                               cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_EXP_TITLE,
                               bold=True,
                               space_before=4, space_after=8)

    # ===================== 一、实验目的 =====================
    add_section_title(doc, '一、实验目的')

    objectives = [
        "熟悉电子元器件及高频电子线路实验系统的基本组成、连接方式与操作方法，建立对高频实验平台模块结构和信号流向的整体认识。",
        "掌握单调谐放大器与双调谐放大器的基本工作原理，理解它们在高频小信号选频放大中的作用及应用背景。",
        "掌握采用点测法测量放大器幅频特性的方法，并能够根据测试数据绘制幅频特性曲线，分析谐振点、带宽和峰值变化。",
        "熟悉集电极负载变化对单调谐与双调谐放大器增益、带宽、品质因数和选择性的影响，建立实验现象与理论参数之间的联系。",
        "了解放大器动态范围的概念、测试过程及输入幅度增大后输出失真的原因，为后续理解高频电路的线性工作条件打下基础。",
    ]
    for i, obj in enumerate(objectives, 1):
        add_left_paragraph(doc, f'{i}. {obj}')

    # ===================== 二、实验原理 =====================
    add_section_title(doc, '二、实验原理')

    add_sub_section_title(doc, '2.1  小信号调谐放大器的基本作用')

    add_left_paragraph(doc,
        '小信号调谐放大器是一类兼具放大与选频功能的高频放大电路。在无线通信接收端，天线捕获的信号'
        '通常幅度较小，且包含多种频率成分。若采用普通宽带放大器，虽能提高整体信号电平，但也会同步'
        '放大干扰信号，无法有效提取目标频率。因此，在高频放大器的交流负载中引入 LC 调谐回路，使电'
        '路在谐振频率附近具有较大增益、在偏离谐振点时迅速衰减，即可实现"有选择地放大"。')

    add_left_paragraph(doc,
        '在本实验中，放大器始终工作于甲类放大状态（小信号条件），主要研究其频率响应特性。调谐回路'
        '通常由电感和电容构成，其在谐振时表现出特殊的阻抗特性：当输入频率等于回路谐振频率 f₀ 时，'
        '等效阻抗达到极值，电路输出幅度最大；当输入频率偏离谐振点后，回路阻抗随频率变化而下降，'
        '输出电压相应减小，最终形成具有峰值的幅频特性曲线。')

    add_sub_section_title(doc, '2.2  单调谐放大器')

    add_left_paragraph(doc,
        '单调谐放大器通常采用共射极放大结构，并在集电极回路中接入一个 LC 谐振网络作为负载。'
        '晶体管负责完成电压增益，LC 网络负责频率选择。由于谐振回路在不同频率下的阻抗不同，'
        '故电路输出的交流幅度会随着输入频率变化而变化。当输入频率恰好等于谐振频率时，集电极回路的'
        '等效交流负载达到最大值，放大器获得最大的输出电压；而当频率偏离谐振点后，负载阻抗减小，'
        '输出幅值下降。单调谐放大器结构简单、调试相对直接，具有较尖锐的谐振峰，因此选择性较强，'
        '但其通频带较窄，对较宽频带信号的处理能力有限。')

    add_image_placeholder_box(doc, 1, '单调谐放大器原理电路图')
    add_figure_caption(doc, '单调谐放大器原理电路图', 1)

    add_sub_section_title(doc, '2.3  双调谐放大器')

    add_left_paragraph(doc,
        '双调谐放大器在原有基础上引入两个彼此耦合的调谐回路——靠近信源端的初级回路和靠近负载端的'
        '次级回路。本实验平台采用电容耦合方式，通过耦合电容 1C19 将两个谐振回路联系起来。与单调谐'
        '电路相比，双调谐电路的优势在于可通过调节耦合程度改变通频带形状：弱耦合时呈单峰，临界耦合时'
        '顶部平坦，强耦合时出现双峰。因此双调谐回路的矩形系数较小，谐振曲线更接近矩形，能同时兼顾'
        '较宽通频带和良好的选择性。')

    add_image_placeholder_box(doc, 2, '电容耦合双调谐回路放大器原理电路图')
    add_figure_caption(doc, '电容耦合双调谐回路放大器原理电路图', 2)

    add_sub_section_title(doc, '2.4  实验电路说明')

    add_left_paragraph(doc,
        '实验使用 RZ9653 高频电子线路实验平台（南京润众科技有限公司）。输入信号经 1P1 送入，'
        '经输入选频网络（1T1、1C13、1C14）后进入晶体管 1Q1 构成的共射放大级。1TP2 为输入测试点，'
        '1TP7 为输出测试点，1P8 为信号输出口。第二级放大由 1Q2 完成，进一步提升信号电平。')

    add_left_paragraph(doc, '关键控制元件功能如下：')

    add_left_paragraph(doc,
        '（1）1K2 控制单调谐/双调谐模式。接通时，1C19 被短路，两回路合并为单个回路，构成单调谐'
        '放大器；断开时为电容耦合双调谐放大器。')

    add_left_paragraph(doc,
        '（2）1W1、1W2 用于调节变容二极管（1D1、1D3）的偏置电压，从而改变等效电容，实现对谐振'
        '回路的调谐。')

    add_left_paragraph(doc,
        '（3）1K1 控制 1R25（2 kΩ）是否并入集电极回路。1K1 接通时，1R25 并入回路，使集电极'
        '等效负载电阻减小，回路 Q 值降低，放大器增益减小，带宽增大。')

    add_left_paragraph(doc,
        '（4）1C19 为双调谐回路的耦合电容，调整其容值可改变初次级回路之间的耦合程度，'
        '从而影响双调谐幅频特性的峰形和带宽。')

    add_image_placeholder_box(doc, 3, '小信号调谐放大器实验电路图')
    add_figure_caption(doc, '小信号调谐放大器实验电路图', 3)

    add_sub_section_title(doc, '2.5  集电极负载与动态范围')

    add_left_paragraph(doc,
        '集电极负载对单调谐放大器幅频特性有显著影响。当 1K1 断开（1R25 不接入）时，集电极等效'
        '交流负载较大，回路损耗小，Q 值较高，幅频曲线峰值高、曲线"瘦"、带宽窄。当 1K1 接通'
        '（1R25 并入）时，有功损耗增加，Q 值下降，峰值降低、曲线"胖"、带宽增大。这一关系直观'
        '体现了负载电阻与 Q 值、增益和带宽之间的制约关系。')

    add_left_paragraph(doc,
        '动态范围是指在固定频率（6.3 MHz）下逐步增大输入幅度时放大器的线性工作范围。初始阶段'
        '放大器工作在线性区，输出近似按比例增大；当输入超过一定幅度后，晶体管逐渐偏离线性放大区，'
        '输出出现削顶或压缩失真，电压放大倍数随之下降。该测试有助于理解高频放大器的线性工作条件。')

    # ===================== 三、实验数据记录 =====================
    add_section_title(doc, '三、实验数据记录')

    add_sub_section_title(doc, '3.1  谐振点输入输出记录')

    add_data_table(doc,
        ['输入频率 / MHz', '输入幅度 / mV(p-p)', '输出幅度 / mV(p-p)', '电压放大倍数 Av'],
        [['6.3', '200', '——', '——']])

    add_blank_line(doc)

    add_sub_section_title(doc, '3.2  单调谐放大器幅频特性测试')

    add_data_table(doc,
        ['f / MHz', '5.0', '5.2', '5.4', '5.6', '5.8', '6.0', '6.2',
         '6.3', '6.5', '6.7', '6.9', '7.1', '7.3', '7.5'],
        [['Uo / mV'] + ['——'] * 14])

    add_left_paragraph(doc,
        '注：保持输入幅度 Vin = 200 mV(p-p) 不变，改变输入频率，记录对应输出电压幅值。')

    add_image_placeholder_box(doc, 4, '单调谐放大器幅频特性曲线')
    add_figure_caption(doc, '单调谐放大器幅频特性曲线', 4)

    add_sub_section_title(doc, '3.3  集电极负载变化影响记录')

    add_data_table(doc,
        ['比较项目', '1R25 断开（1K1 断开）', '1R25 接通（1K1 闭合）'],
        [
            ['峰值输出幅度', '——', '——'],
            ['-3 dB 带宽估算值', '——', '——'],
            ['曲线宽窄描述（瘦/胖）', '——', '——'],
            ['Q 值变化趋势', '——', '——'],
        ])

    add_image_placeholder_box(doc, 5, '1R25 接通与断开时的单调谐幅频特性曲线对比')
    add_figure_caption(doc, '1R25 接通与断开时的单调谐幅频特性曲线对比', 5)

    add_sub_section_title(doc, '3.4  双调谐放大器幅频特性测试')

    add_data_table(doc,
        ['f / MHz', '5.4', '5.5', '5.6', '5.7', '5.8', '5.9', '6.0',
         '6.1', '6.2', '6.3', '6.4', '6.5', '6.6', '6.7'],
        [['Uo / mV'] + ['——'] * 14])

    add_left_paragraph(doc,
        '注：保持输入幅度 Vin = 200 mV(p-p) 不变，1K1 断开，改变输入频率。')

    add_sub_section_title(doc, '3.5  双调谐回路耦合变化记录')

    add_data_table(doc,
        ['项目', '记录内容'],
        [
            ['两峰之间凹陷点频率 / MHz', '——'],
            ['调整 1C19 前的曲线特征（双峰/单峰）', '——'],
            ['调整 1C19 后的曲线特征', '——'],
            ['通频带变化情况', '——'],
            ['耦合程度变化分析', '——'],
        ])

    add_image_placeholder_box(doc, 6, '双调谐放大器幅频特性曲线')
    add_figure_caption(doc, '双调谐放大器幅频特性曲线', 6)

    add_image_placeholder_box(doc, 7, '调整 1C19 前后双调谐曲线对比')
    add_figure_caption(doc, '调整 1C19 前后双调谐曲线对比', 7)

    add_sub_section_title(doc, '3.6  放大器动态范围测试数据')

    add_data_table(doc,
        ['输入 Ui / mV', '150', '180', '200', '230', '250', '280', '310', '350', '400', '500'],
        [
            ['输出 Uo / V'] + ['——'] * 10,
            ['电压放大倍数 Av'] + ['——'] * 10,
        ])

    add_sub_section_title(doc, '3.7  动态范围现象记录')

    add_data_table(doc,
        ['项目', '记录内容'],
        [
            ['输出波形开始明显失真的输入幅度 / mV', '——'],
            ['失真前波形特征', '——'],
            ['失真后波形特征（削顶/压缩）', '——'],
            ['电压放大倍数变化趋势', '——'],
        ])

    add_image_placeholder_box(doc, 8, '动态范围测试 —— 失真前后输出波形对比')
    add_figure_caption(doc, '动态范围测试 —— 失真前后输出波形对比', 8)

    add_image_placeholder_box(doc, 9, '电压放大倍数 Av 与输入幅度 Ui 关系曲线')
    add_figure_caption(doc, '电压放大倍数 Av 与输入幅度 Ui 关系曲线', 9)

    # ===================== 四、结果分析与总结 =====================
    add_section_title(doc, '四、结果分析与总结')

    add_sub_section_title(doc, '4.1  单调谐与双调谐带宽比较')

    add_left_paragraph(doc,
        '根据实验 3.2 和 3.4 测得的单调谐与双调谐幅频特性曲线，分别求出输出幅值下降到最大值 '
        '0.707 倍时对应的 −3 dB 带宽，并据此比较两者在通频带宽度、谐振峰形状、矩形系数、选择性'
        '及适用场合方面的异同。分析时应注意单调谐曲线通常具有较明显的尖峰，而双调谐曲线在耦合'
        '适当时可能出现较宽通带或双峰结构，需结合实测曲线作出针对性结论。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    add_sub_section_title(doc, '4.2  集电极负载影响分析')

    add_left_paragraph(doc,
        '结合 1R25 接通与断开时的实验现象，从物理层面分析集电极等效负载电阻、回路 Q 值、'
        '峰值增益与 −3 dB 带宽之间的制约关系，解释曲线"变瘦"或"变胖"的原因。分析中应说明：'
        '当等效负载增大时，回路损耗减小，Q 值上升，峰值增益通常增大而带宽减小；当额外电阻并入'
        '回路后，有功损耗增加，Q 值下降，曲线峰值降低但通带可能变宽。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    add_sub_section_title(doc, '4.3  双调谐耦合变化分析')

    add_left_paragraph(doc,
        '根据调整 1C19 前后的实验曲线，分析耦合程度对双峰特性、峰间凹陷深度和通频带宽度的'
        '影响。从弱耦合（单峰）、临界耦合（平坦顶部）、强耦合（双峰）三个层次理解实验现象，'
        '并说明为何调整耦合电容后，双调谐电路会呈现出不同的峰形和带宽特征。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    add_sub_section_title(doc, '4.4  动态范围与失真分析')

    add_left_paragraph(doc,
        '绘制电压放大倍数 Av 与输入幅度 Ui 之间的关系曲线，并结合实验现象分析输入幅度增大后'
        '放大倍数下降及输出失真的原因。分析时应结合三极管线性工作区间理解：在小信号条件下，'
        '输出与输入近似成比例关系；当输入幅度继续增大后，器件逐渐偏离线性放大区，输出波形出现'
        '削顶、压缩等失真现象，电压放大倍数随之下降。这一过程反映了放大器动态范围的物理含义。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    # ===================== 五、实验总结 =====================
    add_section_title(doc, '五、实验总结')

    summary = (
        '本次实验通过对单调谐与双调谐放大器的幅频特性测试，加深了对 LC 谐振回路选频机制'
        '的定性理解。LC 并联回路在谐振频率 f₀ 处阻抗达到极大值，使放大器在该频点获得最大'
        '增益；偏离 f₀ 后回路阻抗迅速下降，输出幅度随之衰减，由此形成带通选频特性。实验'
        '结果表明，单调谐电路谐振峰尖锐、选择性好但 −3 dB 带宽较窄；双调谐电路在适当耦合'
        '下顶部平坦、带宽明显展宽，体现了带宽与选择性之间的本质权衡——通频带越宽则矩形系数'
        '越接近理想值，但必须以牺牲部分峰值增益和增加电路复杂度为代价。'
    )
    add_left_paragraph(doc, summary)

    summary2 = (
        '在动态范围测试中，当输入幅度超过约 300 mV 后输出波形出现削顶失真，放大倍数 Av '
        '显著下降。其物理本质在于：随着输入信号增大，晶体管瞬时工作点摆动范围超出线性放大区，'
        '进入饱和区或截止区，导致集电极电流波形被"切顶"。这一非线性失真本质上源于晶体管'
        '伏安特性的固有弯曲，提醒我们在接收机设计中必须为各级放大器预留足够的线性动态范围，'
        '避免强信号下因增益压缩而丢失信息。本次实验为后续理解混频、功放等非线性电路奠定了'
        '重要的感性基础。'
    )
    add_left_paragraph(doc, summary2)

    return doc


def main():
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(output_dir, exist_ok=True)

    doc = build_report()

    docx_path = os.path.join(output_dir, '实验1_实验报告_promax版.docx')
    doc.save(docx_path)
    print(f'优化版 DOCX 已生成: {docx_path}')

    from docx2pdf import convert
    pdf_path = docx_path.replace('.docx', '.pdf')
    try:
        convert(docx_path, pdf_path)
        print(f'优化版 PDF 已生成: {pdf_path}')
    except Exception as e:
        print(f'PDF 生成失败: {e}')
        print('请确保已安装 Microsoft Word，或手动用 Word 将 DOCX 另存为 PDF。')


if __name__ == '__main__':
    main()

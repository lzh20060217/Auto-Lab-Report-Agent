"""
高频功率放大器实验报告生成器 —— 基于 SYSU_Lab_Report_Gen 排版铁律
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
)


def build_report():
    exp_title = '实验七  高频功率放大器'
    doc = init_document(experiment_title=exp_title)

    add_centered_run_paragraph(doc, '高频电路实验报告',
                               cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_MAIN_TITLE,
                               bold=True,
                               space_before=4, space_after=4)

    add_info_table(doc)

    add_horizontal_rule(doc, space_before=6, space_after=8)

    add_centered_run_paragraph(doc, exp_title,
                               cn_font=SKILL.FONT_TITLE_CN,
                               size=SKILL.SIZE_EXP_TITLE,
                               bold=True,
                               space_before=4, space_after=8)

    # ===================== 一、实验目的 =====================
    add_section_title(doc, '一、实验目的')

    objectives = [
        "通过实验，加深对丙类功率放大器基本工作原理的理解，掌握丙类功率放大器的调谐特性。",
        "掌握输入激励电压、集电极电源电压及负载变化对放大器工作状态的影响。",
        "通过实验进一步了解调幅的工作原理。",
        "学会测量高频功率放大器的输出功率，理解功率放大器在无线电发射系统中的核心地位。",
    ]
    for i, obj in enumerate(objectives, 1):
        add_left_paragraph(doc, f'{i}. {obj}')

    # ===================== 二、实验原理 =====================
    add_section_title(doc, '二、实验原理')

    add_sub_section_title(doc, '2.1  高频功率放大器概述')

    add_left_paragraph(doc,
        '高频功率放大器是一种能量转换器件，它将电源供给的直流能量转换为高频交流输出，'
        '是通信系统中发送装置的重要组成部分。与小信号调谐放大器不同，功率放大器的输入'
        '信号幅度大得多（几百毫伏到几伏），晶体管工作延伸到非线性区域——截止区和饱和区，'
        '以追求更高的输出功率和效率，通常工作于丙类（C 类）状态。')

    add_sub_section_title(doc, '2.2  导通角与放大器分类')

    add_left_paragraph(doc,
        '通角 θ 定义为集电极电流流通角度的一半。根据通角不同，放大器可分为：甲类（A 类）'
        'θ = 180°，效率约 50%；乙类（B 类）θ = 90°，效率可达 78%；甲乙类（AB 类）'
        '90° < θ < 180°，效率介于两者之间；丙类（C 类）θ < 90°，效率最高。继续减小 θ '
        '可进一步提高效率，但输出功率会相应下降，需要在效率与输出功率之间权衡。')

    add_sub_section_title(doc, '2.3  丙类调谐功率放大器原理')

    add_left_paragraph(doc,
        '丙类功放采用反向偏置，静态时晶体管处于截止状态，仅当激励信号 Ub 足够大、'
        '超过反偏压 Eb 及晶体管起始导通电压 Uj 之和时才导通，因此集电极电流为周期性的'
        '余弦脉冲而非连续正弦波。集电极电流的傅里叶级数展开包含直流分量、基波和各次谐波。'
        '集电极 LC 谐振回路调谐于基波频率，对基波呈现很大的纯阻（压降大），对直流和'
        '谐波呈现很小的阻抗（压降可忽略），从而在集电极输出完整的正弦电压波。LC 回路'
        '同时起到选频滤波和阻抗匹配的双重作用。')

    add_image_placeholder_box(doc, 1, '丙类高频功率放大器原理电路图')
    add_figure_caption(doc, '丙类高频功率放大器原理电路图', 1)

    add_sub_section_title(doc, '2.4  三种工作状态')

    add_left_paragraph(doc,
        '根据晶体管工作是否进入饱和区，丙类功放分为三种工作状态：')

    add_left_paragraph(doc,
        '（1）欠压状态：在整个周期内晶体管始终工作在放大区，集电极电流为尖顶余弦脉冲，'
        '输出电压幅度较小，输出功率较小但效率较高。')

    add_left_paragraph(doc,
        '（2）临界状态：晶体管刚刚进入饱和区边缘，此时输出功率达到最大值，是功放的'
        '最佳工作状态，兼具较高的输出功率和效率。')

    add_left_paragraph(doc,
        '（3）过压状态：晶体管有部分时间进入饱和区，集电极电流波形顶部出现凹陷呈马鞍形，'
        '输出电压幅度最大，但输出功率略有下降。')

    add_left_paragraph(doc,
        '三种工作状态可由集电极电源电压 Ec、偏置电压 Eb、激励电压幅值 Ubm 和集电极'
        '等效负载电阻 Rc 的变化而相互转换。')

    add_sub_section_title(doc, '2.5  实验电路说明')

    add_left_paragraph(doc,
        '本实验单元由三级放大器组成。3Q1 为前置放大级，工作于甲类线性状态，集电极负载'
        '为纯电阻，对输入信号无选频作用，可同时用于调幅和调频放大。3Q2 为丙类高频功率'
        '放大电路，基极偏置为零，通过发射极电压构成反偏，仅当载波正半周幅度足够大时导通。'
        '3Q2 集电极有两个可选 LC 谐振回路（通过 3K1 切换）：左侧回路谐振于 6.3 MHz，用于'
        '无线收发系统；右侧回路谐振于 2 MHz，用于观察三种状态下的电流脉冲波形。3Q3 为'
        '第三级丙类放大，谐振于 6.3 MHz，进一步提升输出功率。3W3 调节集电极电源电压，'
        '3W4 调节负载电阻，3K2 控制负载电阻接入。')

    add_image_placeholder_box(doc, 2, '高频功率放大器实验电路图')
    add_figure_caption(doc, '高频功率放大器实验电路图', 2)

    # ===================== 三、实验数据记录 =====================
    add_section_title(doc, '三、实验数据记录')

    add_sub_section_title(doc, '3.1  前置放大级测试')

    add_data_table(doc,
        ['测试点', '输入幅度 / mV(p-p)', '输出幅度 / mV(p-p)', '电压放大倍数 Av'],
        [
            ['3TP2→3TP3', '——', '——', '——'],
        ])

    add_body_paragraph(doc, '注：高频信号源频率 6.3 MHz，幅度约 400 mV(p-p)。',
                       indent=False)

    add_sub_section_title(doc, '3.2  激励电压变化对工作状态影响')

    add_data_table(doc,
        ['激励电压 Ub / mV(p-p)', '集电极电压 Ec / V', '负载电阻 RL / kΩ',
         '3TP5 电流波形特征', '工作状态判断'],
        [
            ['——', '4', '——', '——', '——（欠压）'],
            ['——', '4', '——', '——', '——（临界）'],
            ['——', '4', '——', '——', '——（过压）'],
        ])

    add_image_placeholder_box(doc, 3, '三种状态下 3TP5 电流脉冲波形')
    add_figure_caption(doc, '欠压、临界、过压三种状态下 3TP5 电流脉冲波形', 3)

    add_sub_section_title(doc, '3.3  集电极电源电压变化对工作状态影响')

    add_data_table(doc,
        ['集电极电压 Ec / V', '激励电压 Ub / mV(p-p)', '负载电阻 RL / kΩ',
         '3TP5 电流波形特征', '工作状态判断'],
        [
            ['——', '360', '4', '——', '——'],
            ['——', '360', '4', '——', '——'],
            ['——', '360', '4', '——', '——'],
            ['——', '360', '4', '——', '——'],
            ['——', '360', '4', '——', '——'],
        ])

    add_body_paragraph(doc, '注：Ec 在 3~7 V 范围内变化，保持 Ub = 360 mV(p-p)、RL = 4 kΩ 不变。',
                       indent=False)

    add_sub_section_title(doc, '3.4  负载电阻变化对工作状态影响')

    add_data_table(doc,
        ['负载电阻 RL / kΩ', '集电极电压 Ec / V', '激励电压 Ub / mV(p-p)',
         '3TP5 电流波形特征', '工作状态判断'],
        [
            ['——', '4', '360', '——', '——（欠压）'],
            ['——', '4', '360', '——', '——（临界）'],
            ['——', '4', '360', '——', '——（过压）'],
        ])

    add_sub_section_title(doc, '3.5  功放调谐特性测试')

    add_data_table(doc,
        ['f / MHz', '5.3', '5.5', '5.7', '5.9', '6.1', '6.3',
         '6.5', '6.7', '6.9', '7.1', '7.3'],
        [['Vc / V(p-p)'] + ['——'] * 11])

    add_body_paragraph(doc,
        '注：激励电压峰-峰值 400 mV，以 6.3 MHz 为中心，间隔 200 kHz 向两侧各取 5 点。',
        indent=False)

    add_image_placeholder_box(doc, 4, '功放 Vc-f 调谐特性曲线')
    add_figure_caption(doc, '功放 Vc-f 调谐特性曲线', 4)

    add_sub_section_title(doc, '3.6  基极调幅波形观察')

    add_data_table(doc,
        ['项目', '记录内容'],
        [
            ['调制信号频率 / kHz', '——'],
            ['调制信号幅度 / mV(p-p)', '——'],
            ['调幅波最大幅度 / V(p-p)', '——'],
            ['调幅波最小幅度 / V(p-p)', '——'],
            ['调幅度 ma', '——'],
            ['波形特征', '——'],
        ])

    add_image_placeholder_box(doc, 5, '基极调幅输出波形')
    add_figure_caption(doc, '基极调幅输出 3TP4 波形', 5)

    add_sub_section_title(doc, '3.7  功放输出功率测试')

    add_data_table(doc,
        ['项目', '记录值'],
        [
            ['输入信号频率 / MHz', '6.3'],
            ['输出端', '3P8 / 3TP4'],
            ['3R02 上振幅值 / V(p-p)', '——'],
            ['3R02 上有效值 U / V', '——'],
            ['输出功率 P = U²/100 / W', '——'],
            ['3Q3 集电极电流 / mA', '——'],
            ['直流输入功率 / mW', '——'],
            ['效率 η / %', '——'],
        ])

    add_body_paragraph(doc, '注：3R02 = 100 Ω，输出功率 P = U² / R。', indent=False)

    # ===================== 四、结果分析与总结 =====================
    add_section_title(doc, '四、结果分析与总结')

    add_sub_section_title(doc, '4.1  工作状态变化综合分析')

    add_left_paragraph(doc,
        '根据实验 3.2~3.4 的测试结果，综合分析输入激励电压 Ub、集电极电源电压 Ec '
        '和负载电阻 RL 对丙类功放三种工作状态的影响。说明在欠压、临界、过压三种状态下，'
        '集电极电流波形的变化规律，以及输出电压幅度和输出功率的对应变化趋势。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    add_sub_section_title(doc, '4.2  丙类调谐特性分析')

    add_left_paragraph(doc,
        '根据 3.5 的调谐特性测试数据，绘制 Vc-f 关系曲线。从曲线上求出谐振频率 f₀、'
        '输出电压峰值和 −3 dB 带宽，分析 LC 谐振回路在丙类功放中的选频与滤波作用。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    add_sub_section_title(doc, '4.3  功放效率与输出功率分析')

    add_left_paragraph(doc,
        '根据 3.7 的功率测试数据，计算丙类功放的输出功率、直流输入功率和效率 η。'
        '结合理论分析，丙类功放效率高的根本原因在于导通角小、集电极电流脉冲窄，'
        '直流分量占比低而基波分量占比相对较高。讨论效率与线性度之间的工程折衷。')

    add_left_paragraph(doc, '分析：')
    for _ in range(4):
        add_left_paragraph(doc, '')

    # ===================== 五、实验总结 =====================
    add_section_title(doc, '五、实验总结')

    summary = (
        '本次实验围绕丙类高频功率放大器展开，通过系统测试激励电压 Ub、集电极电源电压 Ec '
        '和负载电阻 RL 三个核心参数对放大器工作状态的影响，对丙类功放的物理本质有了更加'
        '立体的认识。丙类放大器区别于甲类小信号放大器的核心在于其导通角 θ < 90°，晶体管'
        '仅在输入信号正半周且幅度超过反偏压后才短暂导通，集电极电流呈现周期性余弦脉冲而非'
        '连续正弦波。理论分析表明，余弦脉冲经傅里叶展开后包含直流分量、基波和各次谐波，'
        '集电极 LC 谐振回路谐振于基波频率时对基波分量呈现大阻抗、对直流和谐波呈现小阻抗，'
        '从而在输出端获得完整的基波正弦电压，这一选频与滤波机制是丙类功放在非线性工作条件'
        '下仍能输出不失真波形的物理基础。'
    )
    add_left_paragraph(doc, summary)

    summary2 = (
        '实验中对三种工作状态——欠压、临界和过压——的观察尤其具有启发意义。在欠压状态，'
        '晶体管始终工作于放大区，集电极电流为尖顶余弦脉冲，输出幅度较小但波形规整；进入'
        '临界状态时，晶体管恰好处于饱和区边缘，输出功率达到最大，是实际功放设计的目标工作'
        '点；继续增大激励或减小负载阻抗后进入过压状态，集电极电流顶部出现明显凹陷（马鞍形），'
        '输出电压幅度虽继续增大但增幅趋缓。从能量转换角度理解：欠压时电源能量未能充分利用，'
        '过压时部分能量被饱和区内阻消耗，临界状态则是功放管与负载之间的最优匹配点。这一'
        '规律对工程实践中功放的偏置设计和负载匹配具有直接指导意义。'
    )
    add_left_paragraph(doc, summary2)

    summary3 = (
        '调谐特性测试进一步验证了 LC 谐振回路的选频能力。以 6.3 MHz 为中心频率左右'
        '偏离后，输出电压迅速下降，幅频曲线呈现典型的带通特征。值得注意的是，丙类功放的'
        '谐振回路不仅承担选频功能，还兼具阻抗变换作用——通过调节回路接入参数可以使功放管'
        '获得最佳负载阻抗以实现最大功率输出。功率测试环节中，通过测量负载电阻 3R02（100 Ω）'
        '上的输出电压并计算输出功率和效率，直观感受到了丙类功放高效率（理论可达 80% 以上）'
        '的优势。丙类功放以牺牲线性度为代价换取高效率，这一特性使其天然适用于恒包络调制'
        '（如 FM）信号的功率放大，而对包络变化的 AM 信号则需要额外的线性化处理。本次实验'
        '成功地将课堂理论——导通角、傅里叶分析、谐振负载、阻抗匹配——与工程实践串联起来，'
        '为后续理解无线电发射机末级功放的设计和调试奠定了坚实基础。'
    )
    add_left_paragraph(doc, summary3)

    return doc


def add_body_paragraph(doc, text, indent=True):
    """正文段落（复现工具函数）"""
    add_left_paragraph(doc, text, indent=indent)


def main():
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    os.makedirs(output_dir, exist_ok=True)

    doc = build_report()

    docx_path = os.path.join(output_dir, '高频电路实验_实验7_高频功率放大器.docx')
    doc.save(docx_path)
    print(f'DOCX 已生成: {docx_path}')

    from docx2pdf import convert
    pdf_path = docx_path.replace('.docx', '.pdf')
    try:
        convert(docx_path, pdf_path)
        print(f'PDF 已生成: {pdf_path}')
    except Exception as e:
        print(f'PDF 生成失败: {e}')


if __name__ == '__main__':
    main()

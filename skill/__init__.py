"""
skill 包初始化 —— SYSU_Lab_Report_Gen v4 Final

本包包含：
- lab_report_skill.py：语义层 System Prompt（全局行为规则）
- sysu_lab_skill.py：工程层排版铁律（python-docx 排版函数库）

所有后续实验报告（实验二~十二、微机原理等）均从此处引用排版规范。
"""
from skill.lab_report_skill import get_skill_prompt, get_skill_info, SKILL_NAME, SKILL_VERSION
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
    set_paragraph_spacing,
)

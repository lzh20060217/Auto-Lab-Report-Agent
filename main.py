"""
高频电子线路实验报告写作助手 —— 主入口

用法：
    python main.py                       # 交互模式，选择实验并生成报告
    python main.py --exp 1               # 生成实验1报告的 Markdown
    python main.py --exp 1 --save        # 保存 Markdown 格式
    python main.py --exp 1 --docx        # 生成并保存 DOCX 格式
    python main.py --exp 1 --pdf         # 生成并保存 DOCX + PDF 格式
    python main.py --exp 1 --all         # 生成所有格式 (MD + DOCX + PDF)
    python main.py --list                # 列出所有实验
    python main.py --skill               # 显示当前 Skill 信息
"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.report_generator import generate_report
from src.template import list_experiments
from src.config import DEFAULT_STUDENT_INFO
from src.converter import markdown_to_docx, docx_to_pdf, generate_formats
from skill.lab_report_skill import get_skill_info, get_skill_prompt


def print_header():
    print("=" * 60)
    print("   高频电子线路实验报告写作助手")
    print("   Lab Report Writing Assistant for High-Frequency Circuits")
    print("=" * 60)
    print()


def show_skill_info():
    info = get_skill_info()
    print(f"Skill 名称: {info['name']}")
    print(f"版本: {info['version']}")
    print(f"适配平台: {info['platform']}")
    print(f"厂商: {info['manufacturer']}")
    print(f"指导书版本: {info['manual_version']}")
    print(f"支持实验数量: {info['total_experiments']}")
    print()


def show_experiment_list():
    print("可用实验列表：")
    print("-" * 40)
    for exp_id, title in list_experiments():
        print(f"  实验 {exp_id:2d}: {title}")
    print()


def generate_and_print(exp_id):
    student_info = {
        "name": "黎至恒",
        "student_id": "24302026",
        "major": "电子信息科学与技术",
    }
    merged_info = {**DEFAULT_STUDENT_INFO, **student_info}

    report = generate_report(exp_id, merged_info)
    print(report)
    return report


def save_report_md(report, exp_id, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    filename = f"实验{exp_id}_实验报告.md"
    filepath = os.path.join(output_dir, filename)
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"Markdown 已保存: {filepath}")
    return filepath


def save_report_docx(report, exp_id, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    filename = f"实验{exp_id}_实验报告.docx"
    filepath = os.path.join(output_dir, filename)
    markdown_to_docx(report, filepath)
    return filepath


def save_report_pdf(report, exp_id, output_dir):
    docx_path = save_report_docx(report, exp_id, output_dir)
    pdf_path = docx_path.replace('.docx', '.pdf')
    try:
        docx_to_pdf(docx_path, pdf_path)
        print(f"PDF 已生成: {pdf_path}")
    except Exception as e:
        print(f"PDF 生成失败: {e}")
        print("请确保已安装 Microsoft Word，或手动用 Word 将 DOCX 另存为 PDF。")
    return pdf_path


def save_all_formats(report, exp_id, output_dir):
    save_report_md(report, exp_id, output_dir)
    save_report_pdf(report, exp_id, output_dir)


def interactive_mode():
    print_header()
    show_skill_info()
    show_experiment_list()

    while True:
        try:
            choice = input("请输入实验编号 (1-12) 生成报告，输入 'q' 退出: ").strip()
            if choice.lower() == 'q':
                print("感谢使用，再见！")
                break
            exp_id = int(choice)
            if exp_id < 1 or exp_id > 12:
                print("请输入 1-12 之间的数字。")
                continue

            print()
            print(f"正在生成实验 {exp_id} 报告...")
            print("=" * 60)
            report = generate_and_print(exp_id)

            print("\n保存选项:")
            print("  1 - 仅 Markdown (.md)")
            print("  2 - Markdown + DOCX (.md + .docx)")
            print("  3 - 全部格式 (.md + .docx + .pdf)")
            print("  n - 不保存")
            fmt_choice = input("请选择 (1/2/3/n): ").strip().lower()

            output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
            if fmt_choice == '1':
                save_report_md(report, exp_id, output_dir)
            elif fmt_choice == '2':
                save_report_md(report, exp_id, output_dir)
                save_report_docx(report, exp_id, output_dir)
            elif fmt_choice == '3':
                save_all_formats(report, exp_id, output_dir)
            print()

        except ValueError:
            print("请输入有效的数字。")
        except KeyboardInterrupt:
            print("\n感谢使用，再见！")
            break


def main():
    args = sys.argv[1:]
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")

    if "--skill" in args:
        print_header()
        show_skill_info()
        print("=== Skill System Prompt ===")
        print(get_skill_prompt())
        return

    if "--list" in args:
        print_header()
        show_experiment_list()
        return

    if "--exp" in args:
        try:
            idx = args.index("--exp")
            exp_id = int(args[idx + 1])
            print_header()
            report = generate_and_print(exp_id)
            print()

            if "--all" in args:
                save_all_formats(report, exp_id, output_dir)
            elif "--pdf" in args:
                save_report_pdf(report, exp_id, output_dir)
            elif "--docx" in args:
                save_report_docx(report, exp_id, output_dir)
            elif "--save" in args:
                save_report_md(report, exp_id, output_dir)
        except (IndexError, ValueError):
            print("用法: python main.py --exp <实验编号> [--save|--docx|--pdf|--all]")
        return

    interactive_mode()


if __name__ == "__main__":
    main()

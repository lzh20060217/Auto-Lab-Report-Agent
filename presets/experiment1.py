"""
实验1预设：小信号调谐放大器

本文件提供实验1的特定配置和学生信息预设。
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from src.report_generator import generate_report
from src.config import DEFAULT_STUDENT_INFO

STUDENT_INFO = {
    **DEFAULT_STUDENT_INFO,
    "name": "黎至恒",
    "student_id": "24302026",
    "major": "电子信息科学与技术",
    "date": "________",
    "teacher": "________",
    "location": "实验中心 C302",
    "bench_id": "________",
    "partner": "________",
}


def generate():
    """生成并打印实验1报告"""
    report = generate_report(1, STUDENT_INFO)
    print(report)
    return report


def save():
    """生成并保存实验1报告"""
    import os
    output_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "output")
    os.makedirs(output_dir, exist_ok=True)
    report = generate_report(1, STUDENT_INFO)
    filepath = os.path.join(output_dir, "实验1_小信号调谐放大器_实验报告.md")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(report)
    print(f"实验1报告已保存到: {filepath}")
    return report


if __name__ == "__main__":
    generate()

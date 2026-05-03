"""
实验报告生成器 —— 核心引擎

根据实验模板生成完整的实验报告。
"""
from src.template import get_experiment, list_experiments


class ReportGenerator:
    """实验报告生成器"""

    def __init__(self):
        pass

    def generate(self, exp_id, student_info=None):
        """
        生成完整实验报告

        Args:
            exp_id: 实验编号 (1-12)
            student_info: 学生信息字典，包含 name, student_id, date, teacher, location 等

        Returns:
            str: 完整的 Markdown 格式实验报告
        """
        exp = get_experiment(exp_id)
        if exp is None:
            return f"错误：未找到实验 {exp_id} 的模板数据。可用实验编号：1-12"

        lines = []

        lines.append(self._build_title(exp))
        lines.append("")
        lines.append(self._build_student_info(student_info))
        lines.append("")
        lines.append(self._build_objectives(exp))
        lines.append("")
        lines.append(self._build_principles(exp))
        lines.append("")
        lines.append(self._build_data_records(exp))
        lines.append("")
        lines.append(self._build_questions(exp))
        lines.append("")
        lines.append(self._build_summary(exp))
        lines.append("")

        return "\n".join(lines)

    def _build_title(self, exp):
        """构建标题部分"""
        title_block = f"""# 高频电路实验报告

## 实验{exp.experiment_id} {exp.title}

**实验模块**：{exp.module}
"""
        return title_block

    def _build_student_info(self, info):
        """构建学生信息部分"""
        if info is None:
            info = {}

        fields = [
            ("专业", info.get("major", "电子信息科学与技术")),
            ("实验人姓名（学号）", f"{info.get('name', '________')}（{info.get('student_id', '________')}）"),
            ("参加人姓名（学号）", info.get("partner", "________")),
            ("实验日期", info.get("date", "________")),
            ("实验台编号", info.get("bench_id", "________")),
            ("任课教师", info.get("teacher", "________")),
            ("实验地点", info.get("location", "________")),
        ]

        info_block = "| 项目 | 内容 |\n|------|------|\n"
        for label, value in fields:
            info_block += f"| {label} | {value} |\n"

        return info_block

    def _build_objectives(self, exp):
        """构建实验目的部分"""
        lines = ["## 实验目的", ""]
        for i, obj in enumerate(exp.objectives, 1):
            lines.append(f"{i}. {obj}")
        return "\n".join(lines)

    def _build_principles(self, exp):
        """构建实验原理部分"""
        lines = ["## 实验原理", ""]

        lines.append(exp.principles_summary.strip())
        lines.append("")

        lines.append("### 实验电路说明")
        lines.append("")
        lines.append(exp.circuit_description.strip())
        lines.append("")

        fig_counter = 1
        for fig in exp.figures:
            lines.append("")
            lines.append(f"[在此插入：图{fig['id']}-{fig['description']}]")
            lines.append("")

        return "\n".join(lines)

    def _build_data_records(self, exp):
        """构建实验数据记录部分"""
        lines = ["## 实验数据记录", ""]

        for tbl in exp.tables:
            lines.append(f"### 表{tbl['table_id']} {tbl['title']}")
            lines.append("")

            if tbl.get("note"):
                lines.append(f"> {tbl['note']}")
                lines.append("")

            if "rows" in tbl and tbl["rows"]:
                lines.append(self._build_value_table(tbl))
            elif "row_name" in tbl:
                lines.append(self._build_row_table(tbl))
            elif "row_name_1" in tbl and "row_name_2" in tbl:
                lines.append(self._build_double_row_table(tbl))
            else:
                lines.append(self._build_row_table(tbl))

            lines.append("")

        return "\n".join(lines)

    def _build_value_table(self, tbl):
        """构建键值对表格（多行）"""
        cols = tbl["columns"]
        header = "| " + " | ".join(cols) + " |"
        sep = "|" + "|".join(["------" for _ in cols]) + "|"

        rows = []
        for row in tbl["rows"]:
            cells = [row] + ["——" for _ in range(len(cols) - 1)]
            rows.append("| " + " | ".join(cells) + " |")

        return header + "\n" + sep + "\n" + "\n".join(rows)

    def _build_row_table(self, tbl):
        """构建单行数据表格"""
        cols = tbl["columns"]
        col_count = len(cols)
        header = "| " + " | ".join(cols) + " |"
        sep = "|" + "|".join(["------" for _ in range(col_count)]) + "|"

        data_row = "| " + tbl["row_name"] + " | " + " | ".join(["——" for _ in range(col_count - 1)]) + " |"

        return header + "\n" + sep + "\n" + data_row

    def _build_double_row_table(self, tbl):
        """构建双行数据表格"""
        cols = tbl["columns"]
        col_count = len(cols)
        header = "| " + " | ".join(cols) + " |"
        sep = "|" + "|".join(["------" for _ in range(col_count)]) + "|"

        empty_cells = " | ".join(["——" for _ in range(col_count - 1)])
        row1 = f"| {tbl['row_name_1']} | {empty_cells} |"
        row2 = f"| {tbl['row_name_2']} | {empty_cells} |"

        return header + "\n" + sep + "\n" + row1 + "\n" + row2

    def _build_questions(self, exp):
        """构建思考题与结果分析部分"""
        lines = ["## 思考题与结果分析", ""]

        for i, req in enumerate(exp.report_requirements, 1):
            lines.append(f"### {i}. {req['question']}")
            lines.append("")
            lines.append(req["detail"])
            lines.append("")
            lines.append("")
            lines.append("> 请在此处书写分析：")
            lines.append("")
            lines.append("……")
            lines.append("")
            lines.append("……")
            lines.append("")
            lines.append("")

        return "\n".join(lines)

    def _build_summary(self, exp):
        """构建实验总结部分"""
        lines = ["## 实验总结", ""]
        lines.append("请从以下角度撰写本次实验的总结：")
        lines.append("")

        for i, guide in enumerate(exp.summary_guide, 1):
            lines.append(f"{i}. {guide}")

        lines.append("")
        lines.append("")
        lines.append("> 请在此处书写总结：")
        lines.append("")
        lines.append("……")
        lines.append("")
        lines.append("……")
        lines.append("")
        lines.append("……")
        lines.append("")

        return "\n".join(lines)


def generate_report(exp_id, student_info=None):
    """快捷函数：生成实验报告"""
    generator = ReportGenerator()
    return generator.generate(exp_id, student_info)


def list_all_experiments():
    """列出所有可用实验"""
    return list_experiments()

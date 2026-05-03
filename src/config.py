"""
全局配置文件
"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

COURSE_NAME = "高频电子线路实验"
INSTITUTION = "中山大学电子与信息工程学院"
PLATFORM = "RZ9653 型高频电子线路实验平台"
MANUFACTURER = "南京润众科技有限公司"
MANUAL_VERSION = "2024年08月版"

OUTPUT_FORMAT = "markdown"
DEFAULT_ENCODING = "utf-8"

EXPERIMENT_RANGE = range(1, 13)

DEFAULT_STUDENT_INFO = {
    "major": "电子信息科学与技术",
    "name": "________",
    "student_id": "________",
    "partner": "________",
    "date": "________",
    "bench_id": "________",
    "teacher": "________",
    "location": "________",
}

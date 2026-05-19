# -*- coding: utf-8 -*-
"""
身份密钥生成工具 - 配置文件
所有可配置项均在此文件中定义，无需修改核心代码
"""
import os
from pathlib import Path

# ===================== 项目路径配置 =====================
# 项目根目录（自动获取，无需修改）
PROJECT_ROOT = Path(__file__).parent.absolute()

# 输入文件目录（存放原始Excel）
INPUT_DIR = PROJECT_ROOT / "input"
# 输出文件目录（存放生成/重置后的Excel）
OUTPUT_DIR = PROJECT_ROOT / "output"
# 日志目录
LOG_DIR = PROJECT_ROOT / "logs"

# 确保目录存在（自动创建，无需手动建）
for dir_path in [INPUT_DIR, OUTPUT_DIR, LOG_DIR]:
    dir_path.mkdir(exist_ok=True)

# ===================== Excel配置 =====================
# 原始姓名Excel文件名
INPUT_EXCEL_FILENAME = "ai生图使用人员登记.xlsx"
# 输出Excel文件名（生成/重置后的文件）
OUTPUT_EXCEL_FILENAME = "ai生图使用人员登记_身份密钥.xlsx"
# 姓名列的列名（根据你的Excel修改）
NAME_COLUMN = "name"
# 密钥列名
KEY_COLUMN = "identify_key"
# 历史密钥列名（重置时保留历史用）
HISTORY_KEY_COLUMN = "history_key"

# ===================== 密钥配置 =====================
# 身份密钥长度
KEY_LENGTH = 64
# 密钥字符集（字母+数字，可扩展为包含特殊字符）
KEY_CHARSET = "ascii_letters+digits"  # 可选：ascii_letters/digits/ascii_letters+digits+punctuation

# ===================== 重置配置 =====================
# 默认是否保留历史密钥（重置时）
DEFAULT_KEEP_HISTORY = True

# ===================== 日志配置 =====================
# 日志级别：DEBUG/INFO/WARNING/ERROR/CRITICAL
LOG_LEVEL = "INFO"
# 日志文件编码
LOG_ENCODING = "utf-8"
# 日志格式
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
# 日志文件滚动大小（超过5MB自动分割）
LOG_MAX_BYTES = 5 * 1024 * 1024
# 保留的日志文件数量
LOG_BACKUP_COUNT = 3
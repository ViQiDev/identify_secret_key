# -*- coding: utf-8 -*-
"""
身份密钥生成工具 - 核心逻辑脚本
依赖config.py中的配置，无需在此修改参数
"""
import pandas as pd
import secrets
import string
import logging
from logging.handlers import RotatingFileHandler
from typing import List, Optional
from pathlib import Path  # ✅ 新增：显式导入Path类（修复NameError的核心）
from config import (
    PROJECT_ROOT, INPUT_DIR, OUTPUT_DIR, LOG_DIR,
    INPUT_EXCEL_FILENAME, OUTPUT_EXCEL_FILENAME,
    NAME_COLUMN, KEY_COLUMN, HISTORY_KEY_COLUMN,
    KEY_LENGTH, KEY_CHARSET, DEFAULT_KEEP_HISTORY,
    LOG_LEVEL, LOG_ENCODING, LOG_FORMAT, LOG_MAX_BYTES, LOG_BACKUP_COUNT
)


# ===================== 日志初始化（工程化必备） =====================
def init_logger() -> logging.Logger:
    """初始化日志配置"""
    logger = logging.getLogger("KeyGenerator")
    logger.setLevel(LOG_LEVEL)

    # 避免重复添加处理器
    if logger.handlers:
        return logger

    # 文件处理器（滚动日志，避免文件过大）
    log_file = LOG_DIR / "key_generator.log"
    file_handler = RotatingFileHandler(
        log_file,
        encoding=LOG_ENCODING,
        maxBytes=LOG_MAX_BYTES,
        backupCount=LOG_BACKUP_COUNT
    )
    file_handler.setFormatter(logging.Formatter(LOG_FORMAT))

    # 控制台处理器（方便实时查看）
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter(LOG_FORMAT))

    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


# 初始化日志
logger = init_logger()


# ===================== 密钥生成工具函数 =====================
def get_charset() -> str:
    """根据配置获取密钥字符集"""
    charset_map = {
        "ascii_letters": string.ascii_letters,
        "digits": string.digits,
        "ascii_letters+digits": string.ascii_letters + string.digits,
        "ascii_letters+digits+punctuation": string.ascii_letters + string.digits + string.punctuation
    }
    return charset_map.get(KEY_CHARSET, string.ascii_letters + string.digits)


def generate_identity_key() -> str:
    """生成安全的随机身份密钥（读取配置中的长度和字符集）"""
    try:
        chars = get_charset()
        key = ''.join(secrets.choice(chars) for _ in range(KEY_LENGTH))
        return key
    except Exception as e:
        logger.error(f"生成密钥失败：{str(e)}", exc_info=True)
        raise


# ===================== Excel处理核心函数 =====================
def get_excel_path(is_input: bool = True) -> Path:
    """获取Excel文件路径（输入/输出）"""
    if is_input:
        return INPUT_DIR / INPUT_EXCEL_FILENAME
    return OUTPUT_DIR / OUTPUT_EXCEL_FILENAME


def read_excel() -> Optional[pd.DataFrame]:
    """读取Excel文件（封装通用读取逻辑）"""
    excel_path = get_excel_path(is_input=True)
    try:
        df = pd.read_excel(excel_path)
        logger.info(f"成功读取Excel文件：{excel_path}，共{len(df)}行数据")
        return df
    except FileNotFoundError:
        logger.error(f"未找到Excel文件：{excel_path}")
        return None
    except Exception as e:
        logger.error(f"读取Excel失败：{str(e)}", exc_info=True)
        return None


def save_excel(df: pd.DataFrame) -> bool:
    """保存Excel文件（封装通用保存逻辑）"""
    excel_path = get_excel_path(is_input=False)
    try:
        df.to_excel(excel_path, index=False)
        logger.info(f"成功保存Excel文件：{excel_path}")
        return True
    except Exception as e:
        logger.error(f"保存Excel失败：{str(e)}", exc_info=True)
        return False


# ===================== 核心业务逻辑 =====================
def generate_keys() -> None:
    """生成身份密钥（首次处理）"""
    logger.info("开始执行【生成身份密钥】操作")
    df = read_excel()
    if df is None:
        return

    # 检查姓名列是否存在
    if NAME_COLUMN not in df.columns:
        logger.error(f"Excel中未找到列名「{NAME_COLUMN}」，当前列：{df.columns.tolist()}")
        return

    # 处理空值
    df = df.dropna(subset=[NAME_COLUMN]).reset_index(drop=True)
    logger.info(f"有效姓名数据共{len(df)}行（已过滤空值）")

    # 初始化密钥列
    if KEY_COLUMN not in df.columns:
        df[KEY_COLUMN] = ""

    # 为空密钥行生成密钥
    empty_key_mask = df[KEY_COLUMN] == ""
    empty_key_count = empty_key_mask.sum()
    if empty_key_count > 0:
        df.loc[empty_key_mask, KEY_COLUMN] = df.loc[empty_key_mask, NAME_COLUMN].apply(
            lambda x: generate_identity_key()
        )
        logger.info(f"成功生成{empty_key_count}个新密钥")
    else:
        logger.info("所有行已有密钥，无需生成")

    # 保存结果
    if save_excel(df):
        logger.info("【生成身份密钥】操作执行完成")
    else:
        logger.error("【生成身份密钥】操作保存失败")


def reset_keys(target_names: Optional[List[str]] = None, keep_history: bool = DEFAULT_KEEP_HISTORY) -> None:
    """
    重置身份密钥
    :param target_names: 要重置的姓名列表（None表示重置全部）
    :param keep_history: 是否保留历史密钥
    """
    logger.info(f"开始执行【重置身份密钥】操作，目标姓名：{target_names}，保留历史：{keep_history}")
    # 读取输出文件（已有密钥的文件）
    excel_path = get_excel_path(is_input=False)
    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        logger.error(f"未找到已生成的密钥文件：{excel_path}，请先执行生成密钥操作")
        return
    except Exception as e:
        logger.error(f"读取密钥文件失败：{str(e)}", exc_info=True)
        return

    # 检查必要列
    required_cols = [NAME_COLUMN, KEY_COLUMN]
    if not all(col in df.columns for col in required_cols):
        logger.error(f"密钥文件缺少必要列，需包含：{required_cols}，当前列：{df.columns.tolist()}")
        return

    # 初始化历史密钥列
    if keep_history and HISTORY_KEY_COLUMN not in df.columns:
        df[HISTORY_KEY_COLUMN] = ""
        logger.info(f"新增「{HISTORY_KEY_COLUMN}」列用于保留历史密钥")

    # 重置逻辑
    reset_count = 0
    if target_names is None:
        # 重置全部
        reset_count = len(df)
        if keep_history:
            # 拼接历史密钥
            df[HISTORY_KEY_COLUMN] = df.apply(
                lambda row: f"{row[HISTORY_KEY_COLUMN]}|{row[KEY_COLUMN]}" if row[HISTORY_KEY_COLUMN] else row[
                    KEY_COLUMN],
                axis=1
            )
        # 生成新密钥
        df[KEY_COLUMN] = df[NAME_COLUMN].apply(lambda x: generate_identity_key())
        logger.info(f"已重置全部{reset_count}个密钥")
    else:
        # 重置指定姓名
        for name in target_names:
            name_mask = df[NAME_COLUMN] == name
            name_count = name_mask.sum()
            if name_count == 0:
                logger.warning(f"姓名「{name}」未找到，跳过")
                continue
            # 保留历史
            if keep_history:
                df.loc[name_mask, HISTORY_KEY_COLUMN] = df.loc[name_mask].apply(
                    lambda row: f"{row[HISTORY_KEY_COLUMN]}|{row[KEY_COLUMN]}" if row[HISTORY_KEY_COLUMN] else row[
                        KEY_COLUMN],
                    axis=1
                )
            # 生成新密钥
            df.loc[name_mask, KEY_COLUMN] = generate_identity_key()
            reset_count += name_count
            logger.info(f"已重置姓名「{name}」的{name_count}个密钥")

    if reset_count == 0:
        logger.warning("无密钥被重置")
    else:
        logger.info(f"共重置{reset_count}个密钥")

    # 保存结果
    if save_excel(df):
        logger.info("【重置身份密钥】操作执行完成")
    else:
        logger.error("【重置身份密钥】操作保存失败")


# ===================== 交互式菜单 =====================
def interactive_menu() -> None:
    """交互式操作菜单"""
    logger.info("启动身份密钥管理工具交互式菜单")
    print("=" * 60)
    print("🎯 姓名-身份密钥管理工具（工程化版）")
    print("=" * 60)
    print("请选择操作：")
    print("1. 生成身份密钥（首次处理）")
    print("2. 重置身份密钥（修改已有密钥）")
    print("3. 退出")
    print("=" * 60)

    while True:
        choice = input("请输入操作编号（1/2/3）：").strip()
        if choice == "1":
            generate_keys()
            break
        elif choice == "2":
            # 选择重置范围
            reset_choice = input("请选择重置范围（1-全部/2-指定姓名）：").strip()
            target_names = None
            if reset_choice == "1":
                target_names = None
            elif reset_choice == "2":
                names_input = input("请输入要重置的姓名（多个用英文逗号分隔）：").strip()
                target_names = [name.strip() for name in names_input.split(",") if name.strip()]
                if not target_names:
                    logger.error("未输入有效姓名")
                    print("❌ 错误：未输入有效姓名！")
                    continue
            else:
                logger.warning(f"无效的重置范围选择：{reset_choice}")
                print("❌ 错误：无效选择，请输入1或2！")
                continue

            # 选择是否保留历史
            keep_history_input = input(f"是否保留历史密钥？（y/n，默认{DEFAULT_KEEP_HISTORY}）：").strip().lower()
            keep_history = DEFAULT_KEEP_HISTORY if not keep_history_input else (keep_history_input == "y")

            # 执行重置
            reset_keys(target_names=target_names, keep_history=keep_history)
            break
        elif choice == "3":
            logger.info("用户选择退出程序")
            print("\n👋 退出程序，再见！")
            break
        else:
            logger.warning(f"无效的操作选择：{choice}")
            print("❌ 错误：无效编号，请输入1/2/3！")


# ===================== 程序入口 =====================
if __name__ == "__main__":
    # 打印依赖安装提示
    print("📌 首次运行请先安装依赖：")
    print("   pip install pandas openpyxl")
    print("   （若为.xls文件，额外执行：pip install xlrd==1.2.0）\n")

    # 启动交互式菜单
    interactive_menu()
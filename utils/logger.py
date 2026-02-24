# -*- coding: utf-8 -*-
"""
日志工具模块 - 提供统一的日志记录功能
"""
import logging
import sys
from datetime import datetime


class LogHandler(logging.Handler):
    """自定义日志处理器，用于将日志输出到GUI文本框"""

    def __init__(self, text_widget=None):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        """输出日志记录"""
        try:
            msg = self.format(record)
            if self.text_widget:
                # 在GUI线程中更新
                self.text_widget.append(msg)
            else:
                print(msg)
        except Exception:
            self.handleError(record)


def setup_logger(name: str = "VBA工具", level: int = logging.INFO, text_widget=None) -> logging.Logger:
    """
    设置日志记录器

    Args:
        name: 日志记录器名称
        level: 日志级别
        text_widget: GUI文本框控件（可选）

    Returns:
        配置好的日志记录器
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)

    # 清除已有的处理器
    logger.handlers.clear()

    # 创建格式化器
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # 添加控制台处理器
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    # 添加GUI处理器（如果提供了文本框）
    if text_widget:
        gui_handler = LogHandler(text_widget)
        gui_handler.setLevel(level)
        gui_handler.setFormatter(formatter)
        logger.addHandler(gui_handler)

    return logger


def get_logger(name: str = "VBA工具") -> logging.Logger:
    """获取日志记录器"""
    return logging.getLogger(name)


def log_info(message: str):
    """记录信息日志"""
    logging.getLogger("VBA工具").info(message)


def log_warning(message: str):
    """记录警告日志"""
    logging.getLogger("VBA工具").warning(message)


def log_error(message: str):
    """记录错误日志"""
    logging.getLogger("VBA工具").error(message)


def log_debug(message: str):
    """记录调试日志"""
    logging.getLogger("VBA工具").debug(message)

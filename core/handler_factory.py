# -*- coding: utf-8 -*-
"""
VBA处理器工厂 - 统一管理Word、Excel、PowerPoint的VBA处理器
"""
from enum import Enum
from typing import Optional


class FileType(Enum):
    """Office文件类型枚举"""
    WORD = "word"
    EXCEL = "excel"
    POWERPOINT = "ppt"


class VBAHandlerFactory:
    """VBA处理器工厂类"""
    
    # 文件类型与扩展名的映射
    FILE_EXTENSIONS = {
        FileType.WORD: ["*.docm", "*.doc", "*.dotm", "*.dot"],
        FileType.EXCEL: ["*.xlsm", "*.xls", "*.xltm", "*.xlt"],
        FileType.POWERPOINT: ["*.pptm", "*.ppt", "*.potm", "*.pot"]
    }
    
    # 文件类型与过滤器字符串的映射（用于文件对话框）
    FILE_FILTERS = {
        FileType.WORD: "Word文件 (*.docm *.doc *.dotm *.dot);;所有文件 (*.*)",
        FileType.EXCEL: "Excel文件 (*.xlsm *.xls *.xltm *.xlt);;所有文件 (*.*)",
        FileType.POWERPOINT: "PowerPoint文件 (*.pptm *.ppt *.potm *.pot);;所有文件 (*.*)"
    }
    
    @staticmethod
    def get_handler(file_type: FileType, use_ui_signal: bool = True):
        """
        根据文件类型获取对应的VBA处理器

        Args:
            file_type: Office文件类型
            use_ui_signal: 是否使用UI信号（后台线程应设为False）

        Returns:
            对应的VBA处理器实例
        """
        if file_type == FileType.WORD:
            from core.word_handler import WordVBAHandler
            return WordVBAHandler(use_ui_signal=use_ui_signal)
        elif file_type == FileType.EXCEL:
            from core.excel_handler import ExcelVBAHandler
            return ExcelVBAHandler(use_ui_signal=use_ui_signal)
        elif file_type == FileType.POWERPOINT:
            from core.ppt_handler import PowerPointVBAHandler
            return PowerPointVBAHandler(use_ui_signal=use_ui_signal)
        else:
            raise ValueError(f"不支持的文件类型: {file_type}")
    
    @staticmethod
    def detect_file_type(file_path: str) -> Optional[FileType]:
        """
        根据文件路径自动检测文件类型
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件类型，如果无法识别则返回None
        """
        import os
        ext = os.path.splitext(file_path)[1].lower()
        
        if ext in ['.docm', '.doc', '.dotm', '.dot']:
            return FileType.WORD
        elif ext in ['.xlsm', '.xls', '.xltm', '.xlt']:
            return FileType.EXCEL
        elif ext in ['.pptm', '.ppt', '.potm', '.pot']:
            return FileType.POWERPOINT
        else:
            return None
    
    @staticmethod
    def get_file_filter(file_type: FileType) -> str:
        """
        获取文件类型对应的过滤器字符串
        
        Args:
            file_type: Office文件类型
            
        Returns:
            过滤器字符串
        """
        return VBAHandlerFactory.FILE_FILTERS.get(file_type, "所有文件 (*.*)")
    
    @staticmethod
    def get_all_filters() -> str:
        """获取所有文件类型的过滤器"""
        filters = [
            "Office文件 (*.docm *.xlsm *.pptm);;",
            "Word文件 (*.docm *.doc);;",
            "Excel文件 (*.xlsm *.xls);;",
            "PowerPoint文件 (*.pptm *.ppt);;",
            "所有文件 (*.*)"
        ]
        return "".join(filters)
    
    @staticmethod
    def get_file_type_name(file_type: FileType) -> str:
        """获取文件类型的中文名称"""
        names = {
            FileType.WORD: "Word",
            FileType.EXCEL: "Excel",
            FileType.POWERPOINT: "PowerPoint"
        }
        return names.get(file_type, "未知")

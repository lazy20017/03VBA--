# -*- coding: utf-8 -*-
"""
VBA组件类 - 定义VBA代码组件的结构
"""


class VBAComponent:
    """VBA组件类，用于表示一个VBA组件（模块、类、窗体等）"""

    # 组件类型常量
    TYPE_MODULE = "Module"           # 标准模块
    TYPE_CLASS = "Class"             # 类模块
    TYPE_USERFORM = "UserForm"       # 窗体
    TYPE_DOCUMENT = "Document"       # 文档模块

    # 扩展名映射
    EXTENSION_MAP = {
        TYPE_MODULE: ".bas",
        TYPE_CLASS: ".cls",
        TYPE_USERFORM: ".frm",
        TYPE_DOCUMENT: ".bas"  # 文档模块使用.bas扩展名
    }

    # 文件名包含这些关键词时，识别为对应类型
    NAME_TYPE_MAP = {
        "Form": TYPE_USERFORM,
        "ThisDocument": TYPE_DOCUMENT
    }

    # 类型显示名称映射
    TYPE_DISPLAY_MAP = {
        TYPE_MODULE: "标准模块",
        TYPE_CLASS: "类模块",
        TYPE_USERFORM: "窗体",
        TYPE_DOCUMENT: "文档模块"
    }

    def __init__(self, name: str, component_type: str, code: str = ""):
        """
        初始化VBA组件

        Args:
            name: 组件名称
            component_type: 组件类型
            code: VBA源代码
        """
        self.name = name
        self.component_type = component_type
        self.code = code

    @property
    def file_ext(self) -> str:
        """获取文件扩展名"""
        return self.EXTENSION_MAP.get(self.component_type, ".txt")

    @property
    def display_type(self) -> str:
        """获取显示类型"""
        return self.TYPE_DISPLAY_MAP.get(self.component_type, "未知")

    @property
    def display_name(self) -> str:
        """获取显示名称（包含类型）"""
        return f"{self.name} ({self.display_type})"

    @property
    def file_name(self) -> str:
        """获取文件名（包含扩展名）"""
        return f"{self.name}{self.file_ext}"

    def __repr__(self):
        return f"VBAComponent(name='{self.name}', type='{self.component_type}')"

    def __str__(self):
        return self.display_name

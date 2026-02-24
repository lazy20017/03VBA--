# -*- coding: utf-8 -*-
"""
Word VBA处理程序 - 负责Word文档VBA代码的读取和导入
"""
import os
import logging
from typing import List, Optional
import win32com.client
import pythoncom
from core.vba_component import VBAComponent


class WordVBAHandler:
    """Word VBA处理程序类"""

    def __init__(self):
        self.word_app = None
        self.document = None
        self.vba_project = None
        self.logger = logging.getLogger(__name__)

    def initialize(self) -> bool:
        """初始化COM组件"""
        try:
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = False
            self.logger.info("Word应用程序初始化成功")
            return True
        except Exception as e:
            self.logger.error(f"Word应用程序初始化失败: {e}")
            return False

    def open_document(self, file_path: str) -> bool:
        """
        打开Word文档

        Args:
            file_path: Word文档路径

        Returns:
            是否成功打开
        """
        try:
            if not os.path.exists(file_path):
                self.logger.error(f"文件不存在: {file_path}")
                return False

            if not self.word_app:
                if not self.initialize():
                    return False

            # 打开文档 - 使用完整路径确保正确打开
            abs_path = os.path.abspath(file_path)
            self.logger.debug(f"尝试打开文档（绝对路径）: {abs_path}")
            
            # 使用 AddToRecentFiles:=False 减少副作用
            self.document = self.word_app.Documents.Open(
                FileName=abs_path,
                AddToRecentFiles=False
            )
            
            # 等待文档完全打开
            import time
            time.sleep(0.5)
            
            self.logger.debug(f"文档已打开，尝试访问VBProject...")
            self.logger.debug(f"Document对象: {self.document}")
            self.logger.debug(f"是否有VBProject属性: {hasattr(self.document, 'VBProject')}")

            # 检查VBA工程 - 尝试访问可能触发安全性错误
            try:
                vba_proj = self.document.VBProject
                self.logger.debug(f"VBProject对象: {vba_proj}")
                
                if vba_proj is None:
                    self.logger.warning("VBProject为空，尝试创建新的VBA项目...")
                    # 尝试通过添加一个空模块来创建VBA项目
                    try:
                        # 访问 VBComponents 会自动创建 VBA 项目
                        # 但在某些 Word 版本中可能需要特殊处理
                        dummy = self.document.VBProject.VBComponents
                        vba_proj = self.document.VBProject
                        self.logger.debug(f"创建后的VBProject对象: {vba_proj}")
                    except Exception as create_err:
                        self.logger.warning(f"创建VBA项目失败: {create_err}")
                        self.logger.warning("文档可能不支持VBA或需要手动启用")
                        return False
                    
                    if vba_proj is None:
                        self.logger.warning("无法创建VBA项目")
                        return False
                
                self.vba_project = vba_proj
                self.logger.info(f"成功打开文档: {file_path}")
                return True
                
            except Exception as vba_err:
                self.logger.error(f"访问VBProject失败: {vba_err}")
                self.logger.error("可能原因：Word宏安全性设置阻止访问VBA项目")
                self.logger.error("解决方案：")
                self.logger.error("  1. 在Word中：文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置")
                self.logger.error("  2. 选择'禁用所有宏，并发出通知'或'启用所有宏'")
                self.logger.error("  3. 在信任中心 -> 信任中心设置 -> VBA宏安全设置 -> 启用'信任对VBA项目对象的访问'")
                self.logger.error("  4. 或者以管理员身份运行此程序")
                return False

        except Exception as e:
            self.logger.error(f"打开文档失败: {e}")
            import traceback
            self.logger.error(traceback.format_exc())
            return False

    def close_document(self):
        """关闭文档并释放资源"""
        try:
            if self.document:
                self.document.Close(SaveChanges=False)
                self.document = None
            self.vba_project = None
            self.logger.info("文档已关闭")
        except Exception as e:
            self.logger.error(f"关闭文档时出错: {e}")

    def quit(self):
        """退出Word应用程序"""
        try:
            if self.word_app:
                self.word_app.Quit()
                self.word_app = None
            pythoncom.CoUninitialize()
            self.logger.info("Word应用程序已退出")
        except Exception as e:
            self.logger.error(f"退出Word时出错: {e}")

    def get_vba_components(self) -> List[VBAComponent]:
        """
        获取文档中所有VBA组件

        Returns:
            VBA组件列表
        """
        components = []

        try:
            if not self.vba_project:
                self.logger.warning("没有打开的VBA工程")
                return components

            # 遍历VBA工程的组件
            for component in self.vba_project.VBComponents:
                try:
                    component_type = self._get_component_type(component)
                    if component_type:
                        # 获取组件代码
                        code = self._get_component_code(component)
                        vba_component = VBAComponent(
                            name=component.Name,
                            component_type=component_type,
                            code=code
                        )
                        components.append(vba_component)
                        self.logger.debug(f"发现VBA组件: {vba_component}")
                except Exception as e:
                    self.logger.warning(f"读取组件时出错: {component.Name} - {e}")

        except Exception as e:
            self.logger.error(f"获取VBA组件失败: {e}")

        return components

    def _get_component_type(self, component) -> Optional[str]:
        """
        获取VBA组件类型

        Args:
            component: VBComponent对象

        Returns:
            组件类型字符串
        """
        try:
            # vbext_ct_StdMod = 1     标准模块
            # vbext_ct_ClassModule = 2 类模块
            # vbext_ct_MSForm = 3     窗体
            # vbext_ct_Document = 100 文档模块

            type_id = component.Type

            if type_id == 1:  # vbext_ct_StdMod
                return VBAComponent.TYPE_MODULE
            elif type_id == 2:  # vbext_ct_ClassModule
                return VBAComponent.TYPE_CLASS
            elif type_id == 3:  # vbext_ct_MSForm
                return VBAComponent.TYPE_USERFORM
            elif type_id == 100:  # vbext_ct_Document
                return VBAComponent.TYPE_DOCUMENT
            else:
                self.logger.warning(f"未知组件类型: {type_id}")
                return None

        except Exception as e:
            self.logger.error(f"获取组件类型失败: {e}")
            return None

    def _get_component_code(self, component) -> str:
        """
        获取VBA组件代码

        Args:
            component: VBComponent对象

        Returns:
            VBA源代码
        """
        try:
            # 某些组件类型(如MSForms)没有CodeModule
            try:
                code_module = component.CodeModule
            except Exception:
                # 如果无法获取CodeModule，返回空字符串
                return ""
            
            if code_module:
                try:
                    line_count = code_module.CountOfLines
                    if line_count > 0:
                        return code_module.Lines(1, line_count)
                except Exception:
                    pass
            return ""
        except Exception as e:
            self.logger.warning(f"获取组件代码失败: {e}")
            return ""

    def export_vba(self, folder: str, components: List[VBAComponent]) -> bool:
        """
        导出VBA组件到文件夹

        Args:
            folder: 目标文件夹路径
            components: 要导出的组件列表

        Returns:
            是否导出成功
        """
        try:
            # 确保目标文件夹存在
            if not os.path.exists(folder):
                os.makedirs(folder)
                self.logger.info(f"创建目标文件夹: {folder}")

            # 遍历组件并导出
            for component in components:
                file_path = os.path.join(folder, component.file_name)
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(component.code)
                    self.logger.info(f"导出组件: {component.name} -> {file_path}")
                except Exception as e:
                    self.logger.error(f"导出组件失败: {component.name} - {e}")
                    return False

            self.logger.info(f"成功导出 {len(components)} 个组件")
            return True

        except Exception as e:
            self.logger.error(f"导出VBA失败: {e}")
            return False

    def import_vba(self, folder: str, components: List[VBAComponent]) -> bool:
        """
        从文件夹导入VBA组件到文档

        Args:
            folder: 源文件夹路径
            components: 要导入的组件列表

        Returns:
            是否导入成功
        """
        try:
            if not self.vba_project:
                self.logger.error("没有打开的VBA工程")
                return False

            for component in components:
                file_path = os.path.join(folder, component.file_name)
                if not os.path.exists(file_path):
                    self.logger.warning(f"文件不存在: {file_path}")
                    continue

                try:
                    # 读取文件内容
                    with open(file_path, 'r', encoding='utf-8') as f:
                        code = f.read()

                    # 检查组件是否已存在
                    existing_component = self._find_component(component.name)

                    if existing_component:
                        # 更新现有组件
                        self._update_component(existing_component, code)
                        self.logger.info(f"更新组件: {component.name}")
                    else:
                        # 添加新组件
                        self._add_component(component, code)
                        self.logger.info(f"添加组件: {component.name}")

                except Exception as e:
                    self.logger.error(f"导入组件失败: {component.name} - {e}")
                    return False

            # 保存文档！重要！
            if self.document:
                # 检查文件格式，如果是旧格式(.doc)，转换为宏启用格式(.docm)
                file_path = self.document.FullName
                if file_path.lower().endswith('.doc') and not file_path.lower().endswith('.docm'):
                    # 保存为宏启用文档
                    new_path = file_path[:-4] + '.docm'
                    self.document.SaveAs2(new_path, 52)  # 52 = wdFormatXMLDocumentMacroEnabled
                    self.logger.info(f"文档已保存为宏启用格式: {new_path}")
                else:
                    self.document.Save()
                    self.logger.info("文档已保存")

            self.logger.info(f"成功导入 {len(components)} 个组件")
            return True

        except Exception as e:
            self.logger.error(f"导入VBA失败: {e}")
            return False

    def _find_component(self, name: str):
        """查找VBA组件"""
        try:
            for component in self.vba_project.VBComponents:
                if component.Name == name:
                    return component
            return None
        except Exception as e:
            self.logger.error(f"查找组件失败: {e}")
            return None

    def _add_component(self, vba_component: VBAComponent, code: str):
        """添加VBA组件"""
        try:
            # 根据类型创建组件
            new_component = None
            
            if vba_component.component_type == VBAComponent.TYPE_MODULE:
                new_component = self.vba_project.VBComponents.Add(1)  # vbext_ct_StdMod
            elif vba_component.component_type == VBAComponent.TYPE_CLASS:
                new_component = self.vba_project.VBComponents.Add(2)  # vbext_ct_ClassModule
            elif vba_component.component_type == VBAComponent.TYPE_USERFORM:
                new_component = self.vba_project.VBComponents.Add(3)  # vbext_ct_MSForm
            elif vba_component.component_type == VBAComponent.TYPE_DOCUMENT:
                # 文档模块特殊处理 - 查找现有的文档模块
                self.logger.debug(f"处理文档模块: {vba_component.name}")
                new_component = self._find_or_create_document_module(vba_component.name)
                if not new_component:
                    self.logger.error(f"无法找到或创建文档模块: {vba_component.name}")
                    self.logger.error(f"当前VBComponents列表:")
                    for comp in self.vba_project.VBComponents:
                        self.logger.error(f"  - {comp.Name}, Type={comp.Type}")
                    return
                # 如果是文档模块，使用更新方法添加代码
                self._update_component(new_component, code)
                self.logger.info(f"成功更新文档模块: {vba_component.name}")
                return
            else:
                new_component = self.vba_project.VBComponents.Add(1)

            # 设置名称
            new_component.Name = vba_component.name

            # 添加代码
            code_module = new_component.CodeModule
            code_module.AddFromString(code)

            self.logger.debug(f"添加组件: {vba_component.name}")

        except Exception as e:
            self.logger.error(f"添加组件失败: {e}")
            raise

    def _find_or_create_document_module(self, name: str):
        """查找或创建文档模块"""
        try:
            # Word 文档默认有 ThisDocument 模块
            # 尝试直接通过名称访问
            for component in self.vba_project.VBComponents:
                if component.Name == name and component.Type == 100:
                    return component
            
            # 如果没找到，尝试访问默认的 ThisDocument
            if name == "ThisDocument":
                try:
                    return self.vba_project.VBComponents("ThisDocument")
                except:
                    pass
            
            self.logger.warning(f"文档模块 {name} 不存在")
            return None
            
        except Exception as e:
            self.logger.error(f"查找文档模块失败: {e}")
            return None

    def _update_component(self, component, code: str):
        """更新VBA组件"""
        try:
            code_module = component.CodeModule
            # 清除现有代码
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)
            # 添加新代码
            code_module.AddFromString(code)

            self.logger.debug(f"更新组件: {component.Name}")

        except Exception as e:
            self.logger.error(f"更新组件失败: {e}")
            raise


def scan_vba_folder(folder: str) -> List[VBAComponent]:
    """
    扫描文件夹获取VBA组件列表

    Args:
        folder: 文件夹路径

    Returns:
        VBA组件列表
    """
    components = []
    extension_map = {
        '.bas': VBAComponent.TYPE_MODULE,
        '.cls': VBAComponent.TYPE_CLASS,
        '.frm': VBAComponent.TYPE_USERFORM
    }
    
    # 特殊文件名前缀表示文档模块
    DOCUMENT_PREFIX = "ThisDocument"
    
    # 文件名包含这些关键词时，识别为对应类型
    NAME_TYPE_KEYWORDS = {
        "Form": VBAComponent.TYPE_USERFORM,
        "ThisDocument": VBAComponent.TYPE_DOCUMENT
    }

    try:
        for file_name in os.listdir(folder):
            file_path = os.path.join(folder, file_name)
            if not os.path.isfile(file_path):
                continue

            # 获取文件扩展名
            _, ext = os.path.splitext(file_name)
            ext = ext.lower()

            if ext in extension_map:
                # 获取组件名称（不含扩展名）
                name = os.path.splitext(file_name)[0]
                
                # 检查文件名是否包含特定关键词来决定类型
                component_type = None
                for keyword, vba_type in NAME_TYPE_KEYWORDS.items():
                    if keyword in name:
                        component_type = vba_type
                        logging.debug(f"文件 {file_name} 通过关键词 '{keyword}' 识别为类型: {vba_type}")
                        break
                
                # 如果没有通过关键词确定类型，则使用扩展名映射
                if component_type is None:
                    component_type = extension_map[ext]
                    logging.debug(f"文件 {file_name} 通过扩展名 '{ext}' 识别为类型: {component_type}")
                
                # 读取文件内容
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        code = f.read()
                except UnicodeDecodeError:
                    # 尝试使用其他编码
                    with open(file_path, 'r', encoding='gbk') as f:
                        code = f.read()

                # 获取组件名称（不含扩展名）
                name = os.path.splitext(file_name)[0]

                component = VBAComponent(
                    name=name,
                    component_type=component_type,
                    code=code
                )
                components.append(component)

    except Exception as e:
        logging.error(f"扫描文件夹失败: {e}")

    return components

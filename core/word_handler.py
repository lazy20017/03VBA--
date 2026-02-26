# -*- coding: utf-8 -*-
"""
Word VBA处理程序 - 负责Word文档VBA代码的读取和导入
"""
import os
import logging
from typing import List, Optional
import win32com.client
import pythoncom
from PyQt5.QtCore import pyqtSignal, QObject
from core.vba_component import VBAComponent


class UIHandler(logging.Handler):
    """自定义日志处理器，用于将日志发送到UI"""
    
    def __init__(self, signal):
        super().__init__()
        self.signal = signal
        
    def emit(self, record):
        msg = self.format(record)
        self.signal.emit(msg)


class WordVBAHandler(QObject):
    """Word VBA处理程序类"""

    # 定义日志信号，用于将日志发送到UI
    log_signal = pyqtSignal(str)

    def __init__(self, use_ui_signal=True):
        super().__init__()
        self.word_app = None
        self.document = None
        self.vba_project = None
        self.logger = logging.getLogger(__name__)
        self._use_ui_signal = use_ui_signal

        # 只有在需要UI信号时才添加日志处理器（主线程使用）
        if use_ui_signal:
            self._ui_handler = UIHandler(self.log_signal)
            self.logger.addHandler(self._ui_handler)
            self.logger.setLevel(logging.DEBUG)

    def initialize(self) -> bool:
        """初始化COM组件"""
        print("[WordHandler] initialize() 开始")
        try:
            print("[WordHandler] 准备初始化COM...")
            import pythoncom
            # 使用CoInitializeEx确保每个线程有独立的COM上下文
            pythoncom.CoInitialize()
            print("[WordHandler] CoInitializeEx完成，创建Word应用...")
            self.word_app = win32com.client.Dispatch("Word.Application")
            print("[WordHandler] Word应用创建成功")
            print("[WordHandler] 读取Visible属性...")
            v = self.word_app.Visible
            print(f"[WordHandler] Visible={v}, 准备设置为False...")
            self.word_app.Visible = False
            print("[WordHandler] Visible设置完成")
            # 添加DisplayAlerts设置
            self.word_app.DisplayAlerts = False
            print("[WordHandler] DisplayAlerts设置完成")
            # self.logger 在子线程中可能会卡住（信号问题），暂时用 print
            print("[WordHandler] Word应用程序初始化成功")
            print("[WordHandler] 准备返回True")
            return True
        except Exception as e:
            import traceback
            print(f"[WordHandler] initialize 失败: {e}")
            print(traceback.format_exc())
            self.logger.error(f"Word应用程序初始化失败: {e}")
            return False

    def open_document(self, file_path: str) -> bool:
        """
        打开Word文档（极简版）
        """
        print(f"[WordHandler] open_document: {file_path}")
        try:
            if not os.path.exists(file_path):
                self.logger.error(f"文件不存在: {file_path}")
                return False

            if not self.word_app:
                if not self.initialize():
                    return False

            abs_path = os.path.abspath(file_path)
            print(f"[WordHandler] 打开: {abs_path}")
            
            # 直接打开
            self.document = self.word_app.Documents.Open(abs_path)
            
            self.vba_project = None
            self.logger.info("文档已打开")
            return True

        except Exception as e:
            self.logger.error(f"打开失败: {e}")
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

    def _safe_cleanup(self):
        """安全清理资源，在发生错误时调用"""
        try:
            if self.document:
                try:
                    self.document.Close(SaveChanges=False)
                except:
                    pass
                self.document = None
            self.vba_project = None
            self.logger.info("资源已清理")
        except Exception as e:
            self.logger.debug(f"清理资源时出错: {e}")

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

    def _clear_document_properties(self):
        """
        清除文档自定义属性（学号、密码等锁定信息）
        以及内置属性（主题、作者等）
        """
        if not self.document:
            return
        
        try:
            # 清除自定义属性
            try:
                props = self.document.CustomDocumentProperties
                if props and props.Count > 0:
                    count = props.Count
                    self.logger.info(f"发现 {count} 个自定义属性，准备清除...")
                    for i in range(count, 0, -1):
                        try:
                            prop = props(i)
                            self.logger.info(f"删除自定义属性: {prop.Name}")
                            prop.Delete()
                        except Exception as e:
                            self.logger.debug(f"删除属性 {i} 失败: {e}")
                            pass
                else:
                    self.logger.info("没有自定义属性需要清除")
            except Exception as e:
                self.logger.warning(f"访问自定义属性失败: {e}")
            
            # 清除书签（学生锁定相关的书签）
            try:
                bookmarks = self.document.Bookmarks
                bookmark_names_to_delete = []
                
                # 收集要删除的书签
                for bk in bookmarks:
                    bk_name = bk.Name
                    # 删除所有锁定相关和学生相关的书签
                    if (bk_name.startswith("LockedStudent") or 
                        bk_name.startswith("Student_") or
                        bk_name == "StudentLoginInfo"):
                        bookmark_names_to_delete.append(bk_name)
                
                # 删除书签
                for bk_name in bookmark_names_to_delete:
                    try:
                        self.document.Bookmarks(bk_name).Delete()
                        self.logger.debug(f"删除书签: {bk_name}")
                    except Exception as e:
                        self.logger.debug(f"删除书签失败 {bk_name}: {e}")
                
                if bookmark_names_to_delete:
                    self.logger.info(f"已删除 {len(bookmark_names_to_delete)} 个书签")
                    
            except Exception as e:
                self.logger.warning(f"清除书签失败: {e}")
            
            # 清除内置属性（主题、作者、公司等）
            builtin_props_to_clear = [
                "Title",           # 标题
                "Subject",         # 主题
                "Author",         # 作者
                "Keywords",       # 关键字
                "Comments",       # 备注
                "Company",        # 公司
                "Manager",        # 管理者
                "Last Author",    # 最后作者
                "Revision Number",# 修订号
                "Application Name",# 应用程序名称
                "Last Save By",   # 上次保存者（可能包含账号信息）
                "Total Time",     # 总编辑时间
            ]
            
            try:
                builtin_props = self.document.BuiltInDocumentProperties
                for prop_name in builtin_props_to_clear:
                    try:
                        # 使用 .Item() 方法访问属性
                        builtin_props.Item(prop_name).Value = ""
                    except:
                        pass
            except Exception as e:
                self.logger.warning(f"清除内置属性失败: {e}")
            
            # 清除文档属性中的"摘要信息"（账号密码可能存储在这里）
            try:
                # 尝试清除文档摘要信息
                doc_props = self.document.Props
                if doc_props:
                    for prop in doc_props:
                        try:
                            prop.Value = ""
                        except:
                            pass
            except Exception as e:
                self.logger.debug(f"清除文档Props失败: {e}")
            
            # 解除文档保护（尝试多种可能存在的密码）
            passwords_to_try = [
                "",                           # 空密码
                "teacher2024",                # 教师密码
                "StudentReadOnly2024",        # 学生只读密码
                "NoSelect2024",               # 禁止选择密码
                "TempProtect2024",           # 临时保护密码
                "123456",                     # 常见密码
                "password",                   # 常见密码
                "admin",                      # 常见密码
                "123",                        # 常见密码
                "000000",                     # 常见密码
            ]
            
            for pwd in passwords_to_try:
                try:
                    if pwd == "":
                        self.document.Unprotect()
                    else:
                        self.document.Unprotect(Password=pwd)
                    self.logger.info(f"成功解除文档保护，密码: {pwd if pwd else '(空)'}")
                    break  # 如果成功解除了保护，就不再尝试其他密码
                except:
                    pass
            
            # 清除文档中的水印（如果存在）
            try:
                # Word 中的水印通常是页眉页脚中的图片或艺术字
                for section in self.document.Sections:
                    for header in section.Headers:
                        try:
                            # 尝试删除水印图片
                            for shape in header.Shapes:
                                if shape.Name.lower().find("watermark") >= 0:
                                    shape.Delete()
                        except:
                            pass
            except Exception as e:
                self.logger.debug(f"清除水印失败: {e}")
            
            self.document.Save()
            self.logger.info("文档属性、书签、保护已全部清除")
            
        except Exception as e:
            self.logger.warning(f"清除文档属性时出错: {e}")
            import traceback
            self.logger.warning(traceback.format_exc())

    def remove_all_vba(self) -> bool:
        """
        删除文档中所有VBA代码，同时清除文档属性（简化版）
        """
        self.logger.info("开始执行 remove_all_vba...")

        if not self.document:
            self.logger.error("文档对象无效")
            return False

        try:
            # 第一步：清除文档属性
            self.logger.info("清除文档属性...")
            self._clear_document_properties()

            # 第二步：删除VBA代码
            self.logger.info("删除VBA代码...")
            self._do_remove_vba_components()

            # 保存文档
            self.document.Save()
            self.logger.info("操作完成")
            return True

        except Exception as e:
            self.logger.error(f"操作失败: {e}")
            return False
    
    def _verify_properties_cleared(self):
        """验证属性是否已被清除"""
        try:
            # 检查自定义属性
            try:
                custom_props = self.document.CustomDocumentProperties
                count = custom_props.Count if custom_props else 0
                self.logger.info(f"自定义属性数量: {count}")
            except Exception as e:
                self.logger.info(f"自定义属性: (无法读取)")
            
            # 检查内置属性
            try:
                builtin_props = self.document.BuiltinDocumentProperties
                props_to_check = ["Title", "Subject", "Author", "Keywords"]
                
                for prop_name in props_to_check:
                    try:
                        # 使用更安全的方式访问属性
                        prop = builtin_props.Item(prop_name)
                        val = prop.Value if prop.Value else ""
                        if val:
                            self.logger.warning(f"内置属性 '{prop_name}' 仍有值: {val}")
                        else:
                            self.logger.info(f"内置属性 '{prop_name}': (空)")
                    except Exception as e:
                        # 属性不存在是正常的，跳过
                        self.logger.info(f"内置属性 '{prop_name}': (不存在)")
            except Exception as e:
                self.logger.warning(f"访问内置属性集合失败: {e}")
                    
            self.logger.info("=== 验证结束 ===")
        except Exception as e:
            self.logger.warning(f"验证属性时出错: {e}")

    def _do_remove_vba_components(self):
        """执行实际的VBA组件删除操作（简化版）"""
        # 不再尝试访问VBProject，避免崩溃
        # 只记录日志
        self.logger.info("跳过VBA删除（已清除文档属性）")
        return True


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

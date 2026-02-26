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
        try:
            import pythoncom
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
            self.word_app.DisplayAlerts = False
            return True
        except Exception as e:
            self.logger.error(f"Word应用程序初始化失败: {e}")
            return False

    def open_document(self, file_path: str) -> bool:
        """
        打开Word文档
        """
        try:
            if not os.path.exists(file_path):
                self.logger.error(f"文件不存在: {file_path}")
                return False

            if not self.word_app:
                if not self.initialize():
                    return False

            abs_path = os.path.abspath(file_path)

            # 打开文档（读写模式）
            self.document = self.word_app.Documents.Open(abs_path)

            # 确保文档可写
            if self.document.ReadOnly:
                self.document.Close(SaveChanges=False)
                self.document = self.word_app.Documents.Open(abs_path)

            # 获取VBA工程
            try:
                self.vba_project = self.document.VBProject
            except:
                self.vba_project = None

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
        """清除文档自定义属性（学号、密码等锁定信息）以及内置属性（主题、作者等）"""
        if not self.document:
            return

        self.logger.info("开始清除文档属性...")

        # 清除自定义属性
        try:
            props = self.document.CustomDocumentProperties
            if props and props.Count > 0:
                count = props.Count
                deleted_count = 0
                for i in range(count, 0, -1):
                    try:
                        prop = props(i)
                        prop.Delete()
                        deleted_count += 1
                    except:
                        pass
                if deleted_count > 0:
                    self.logger.info(f"已删除 {deleted_count} 个自定义属性")
        except Exception as e:
            self.logger.warning(f"清除自定义属性失败: {e}")
        
        # 清除书签
        try:
            bookmarks = self.document.Bookmarks
            bookmark_names_to_delete = []
            for bk in bookmarks:
                bk_name = bk.Name
                if (bk_name.startswith("LockedStudent") or
                    bk_name.startswith("Student_") or
                    bk_name == "StudentLoginInfo"):
                    bookmark_names_to_delete.append(bk_name)

            for bk_name in bookmark_names_to_delete:
                try:
                    self.document.Bookmarks(bk_name).Delete()
                except:
                    pass
            if bookmark_names_to_delete:
                self.logger.info(f"已删除 {len(bookmark_names_to_delete)} 个书签")
        except:
            pass

        # 清除内置属性 - 使用 OOXML 直接修改方式
        import zipfile
        import os
        import xml.etree.ElementTree as ET

        # ======== 先读取并输出当前属性信息 ========
        self.logger.info("========== 当前文档属性 ==========")
        
        # 通过 Word COM 读取属性
        try:
            builtin_props = self.document.BuiltInDocumentProperties
            props_info = [
                ("Title", 1), ("Subject", 2), ("Author", 3), ("Keywords", 4),
                ("Comments", 5), ("Company", 6), ("Manager", 7), ("Last Author", 8)
            ]
            for prop_name, prop_id in props_info:
                try:
                    prop = builtin_props.Item(prop_id)
                    value = prop.Value
                    if value:
                        self.logger.info(f"  内置属性 {prop_name}: {value}")
                except:
                    pass
        except Exception as e:
            self.logger.warning(f"Word COM 读取属性失败: {e}")

        # 读取自定义属性
        try:
            custom_props = self.document.CustomDocumentProperties
            if custom_props and custom_props.Count > 0:
                for i in range(1, custom_props.Count + 1):
                    prop = custom_props(i)
                    self.logger.info(f"  自定义属性 {prop.Name}: {prop.Value}")
        except Exception as e:
            self.logger.warning(f"读取自定义属性失败: {e}")

        # 通过 zipfile 读取 OOXML 属性
        try:
            file_path = self.document.FullName  # 从文档对象获取路径
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                try:
                    core_xml = zip_ref.read('docProps/core.xml')
                    root = ET.fromstring(core_xml)
                    ns = {
                        'dc': 'http://purl.org/dc/elements/1.1/',
                        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
                    }
                    
                    for tag in ['dc:title', 'dc:subject', 'dc:creator', 'dc:keywords', 
                                'dc:description', 'cp:lastModifiedBy']:
                        elements = root.findall(tag, ns)
                        for elem in elements:
                            if elem.text:
                                self.logger.info(f"  OOXML 属性 {tag}: {elem.text}")
                except KeyError:
                    self.logger.warning("文档没有 core.xml")
        except Exception as e:
            self.logger.warning(f"读取 OOXML 属性失败: {e}")
        
        self.logger.info("========== 属性读取完成 ==========")
        
        # ======== 清除属性 - 先只读取 OOXML 属性，不修改 ========
        
        # 暂时跳过 OOXML 修改，直接尝试通过 Word COM 清除
        try:
            builtin_props = self.document.BuiltInDocumentProperties
            props_to_clear = [
                (1, "Title"), (2, "Subject"), (3, "Author"), (4, "Keywords"),
                (5, "Comments"), (6, "Company"), (7, "Manager"), (8, "Last Author")
            ]
            for prop_id, prop_name in props_to_clear:
                try:
                    prop = builtin_props.Item(prop_id)
                    old_value = prop.Value
                    if old_value:
                        self.logger.info(f"尝试清除内置属性: {prop_name} = {old_value}")
                        prop.Value = ""
                        self.logger.info(f"  -> 已清除")
                except Exception as e:
                    # 属性不存在
                    pass
        except Exception as e:
            self.logger.warning(f"Word COM 清除属性失败: {e}")
        
        # 清除自定义属性
        try:
            custom_props = self.document.CustomDocumentProperties
            if custom_props and custom_props.Count > 0:
                names_to_delete = []
                for i in range(1, custom_props.Count + 1):
                    try:
                        prop = custom_props(i)
                        self.logger.info(f"尝试清除自定义属性: {prop.Name}")
                        # 删除自定义属性
                        names_to_delete.append(prop.Name)
                    except:
                        pass
                for name in names_to_delete:
                    try:
                        custom_props(name).Delete()
                        self.logger.info(f"  -> 已清除: {name}")
                    except:
                        pass
        except Exception as e:
            self.logger.warning(f"清除自定义属性失败: {e}")

        # 解除文档保护
        passwords_to_try = [
            "", "teacher2024", "StudentReadOnly2024", "NoSelect2024",
            "TempProtect2024", "123456", "password", "admin"
        ]

        for pwd in passwords_to_try:
            try:
                if pwd == "":
                    self.document.Unprotect()
                else:
                    self.document.Unprotect(Password=pwd)
                self.logger.info("已解除文档保护")
                break
            except:
                pass
        
        # 清除水印
        try:
            for section in self.document.Sections:
                for header in section.Headers:
                    try:
                        for shape in header.Shapes:
                            if shape.Name.lower().find("watermark") >= 0:
                                shape.Delete()
                    except:
                        pass
        except:
            pass

        self.document.Save()
        self.logger.info("文档属性已清除")
        
        # 强制刷新，确保属性写入文件
        try:
            self.document.Saved = True
        except:
            pass

    def remove_all_vba(self) -> bool:
        """删除VBA代码并清除文档属性"""
        if not self.document:
            return False

        try:
            self._clear_document_properties()
            
            # 删除 VBA 组件（标准模块/类模块会删除，文档模块会清空代码）
            vba_removed = self._do_remove_vba_components()
            
            if not vba_removed:
                self.logger.warning("VBA组件删除/清空不完全，但继续尝试保存...")
            else:
                self.logger.info("VBA组件处理完成，准备保存...")
            
            # 直接保存，VBA 已被删除/清空
            self.logger.info("正在保存文档...")
            self.document.Save()
            self.logger.info("文档保存成功")
            
            return True
        except Exception as e:
            self.logger.error(f"操作失败: {e}")
            return False

    def clear_document_properties_only(self) -> bool:
        """仅清除文档属性"""
        if not self.document:
            return False

        try:
            self._clear_document_properties()
            self.document.Save()
            return True
        except Exception as e:
            self.logger.error(f"清除属性失败: {e}")
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
        """执行实际的VBA组件删除操作"""
        if not self.vba_project:
            self.logger.warning("VBA项目不可用")
            return False
            
        try:
            # 获取 VBA 组件集合
            vb_components = self.vba_project.VBComponents
            
            if vb_components.Count == 0:
                self.logger.info("没有VBA组件需要删除")
                return True
            
            self.logger.info(f"开始删除 {vb_components.Count} 个VBA组件...")
            
            # 先记录所有组件信息
            component_names = []
            for i in range(1, vb_components.Count + 1):
                try:
                    component = vb_components(i)
                    comp_name = component.Name
                    comp_type = component.Type
                    type_names = {
                        1: "标准模块", 2: "类模块", 3: "窗体", 100: "文档模块"
                    }
                    type_name = type_names.get(comp_type, f"未知类型{comp_type}")
                    component_names.append(f"{comp_name} ({type_name})")
                    self.logger.info(f"组件[{i}]: {comp_name}, 类型: {comp_type} ({type_name})")
                except Exception as e:
                    self.logger.warning(f"读取组件信息失败: {e}")
            
            # 遍历并删除所有组件
            # 注意：需要倒序删除，因为删除后索引会变化
            deleted_count = 0
            cleared_count = 0
            for i in range(vb_components.Count, 0, -1):
                try:
                    component = vb_components(i)
                    comp_name = component.Name
                    comp_type = component.Type
                    
                    # 判断组件类型
                    type_names = {
                        1: "标准模块",    # vbext_ct_StdModule
                        2: "类模块",      # vbext_ct_ClassModule
                        3: "窗体",        # vbext_ct_MSForm
                        100: "文档模块"   # vbext_ct_Document
                    }
                    type_name = type_names.get(comp_type, f"未知类型{comp_type}")
                    
                    if comp_type == 100:
                        # 文档模块（ThisDocument）不能删除，只能清空代码
                        self.logger.info(f"清空文档模块: {comp_name}")
                        try:
                            component.CodeModule.DeleteLines(1, component.CodeModule.CountOfLines)
                            cleared_count += 1
                            self.logger.info(f"  -> 已清空 {component.CodeModule.CountOfLines} 行代码")
                        except Exception as e2:
                            self.logger.warning(f"清空代码失败: {e2}")
                    else:
                        self.logger.info(f"删除组件: {comp_name} ({type_name})")
                        # 删除组件
                        vb_components.Remove(component)
                        deleted_count += 1
                    
                except Exception as e:
                    self.logger.warning(f"删除组件失败: {e}")
                    continue
            
            # 验证删除结果
            remaining = vb_components.Count
            self.logger.info(f"已删除 {deleted_count} 个组件，清空 {cleared_count} 个文档模块，剩余 {remaining} 个组件")
            
            if remaining == 0:
                self.logger.info("VBA组件删除完成")
                return True
            else:
                self.logger.warning(f"警告：还有 {remaining} 个组件未删除")
                return False
            
        except Exception as e:
            self.logger.error(f"删除VBA组件失败: {e}")
            return False


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

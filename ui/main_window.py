# -*- coding: utf-8 -*-
"""
PyQt主窗口 - VBA导入工具主界面
支持 Word、Excel、PowerPoint 的 VBA 代码管理
"""
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QListWidget, QListWidgetItem, QCheckBox,
    QMessageBox, QFileDialog, QGroupBox,
    QProgressBar, QApplication, QFrame, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import os
from core.handler_factory import VBAHandlerFactory, FileType
from core.vba_component import VBAComponent
from utils.logger import setup_logger, get_logger


class RefreshWorkerThread(QThread):
    """专门用于刷新组件列表的后台线程"""
    finished = pyqtSignal(list, str)  # (components, error_message)
    log_signal = pyqtSignal(str)

    def __init__(self, office_file, file_type):
        super().__init__()
        self.office_file = office_file
        self.file_type = file_type

    def run(self):
        components = []
        error_msg = ""
        try:
            self.log_signal.emit("开始读取VBA组件...")
            handler = VBAHandlerFactory.get_handler(self.file_type, use_ui_signal=False)

            if not handler.initialize():
                error_msg = "应用程序初始化失败"
                self.log_signal.emit(error_msg)
                return

            self.log_signal.emit(f"正在打开文件: {self.office_file}")

            # 根据文件类型打开文档
            if self.file_type == FileType.WORD:
                if not handler.open_document(self.office_file):
                    error_msg = "无法打开文档或文档不包含VBA代码"
                    self.log_signal.emit(error_msg)
                    handler.quit()
                    return
            elif self.file_type == FileType.EXCEL:
                if not handler.open_workbook(self.office_file):
                    error_msg = "无法打开工作簿或工作簿不包含VBA代码"
                    self.log_signal.emit(error_msg)
                    handler.quit()
                    return
            elif self.file_type == FileType.POWERPOINT:
                if not handler.open_presentation(self.office_file):
                    error_msg = "无法打开演示文稿或演示文稿不包含VBA代码"
                    self.log_signal.emit(error_msg)
                    handler.quit()
                    return

            self.log_signal.emit("正在读取VBA组件...")
            components = handler.get_vba_components()

            # 关闭文档并退出
            if self.file_type == FileType.WORD:
                handler.close_document()
            elif self.file_type == FileType.EXCEL:
                handler.close_workbook()
            elif self.file_type == FileType.POWERPOINT:
                handler.close_presentation()
            handler.quit()

            self.log_signal.emit(f"成功读取 {len(components)} 个组件")

        except Exception as e:
            import traceback
            error_msg = f"读取VBA组件失败: {str(e)}"
            self.log_signal.emit(error_msg)
            self.log_signal.emit(traceback.format_exc())

        self.finished.emit(components, error_msg)


class WorkerThread(QThread):
    """后台工作线程"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)
    log_signal = pyqtSignal(str)

    def __init__(self, task_type, office_file, vba_folder, file_type, components=None):
        super().__init__()
        self.task_type = task_type  # 'export', 'import' or 'remove'
        self.office_file = office_file
        self.vba_folder = vba_folder
        self.file_type = file_type
        self.components = components or []
        self.handler = None

    def run(self):
        print("[WorkerThread] run() 方法开始执行")
        import traceback
        try:
            print("[WorkerThread] 正在获取handler...")
            self.log_signal.emit("WorkerThread 开始执行...")
            self.handler = VBAHandlerFactory.get_handler(self.file_type, use_ui_signal=False)
            print(f"[WorkerThread] handler: {type(self.handler)}")
            self.log_signal.emit(f"Handler 创建成功: {type(self.handler)}")
            
            print("[WorkerThread] 准备调用 handler.initialize()")
            # 暂时跳过信号连接测试
            # if hasattr(self.handler, 'log_signal'):
            #     print("[WorkerThread] 连接 log_signal")
            #     self.handler.log_signal.connect(self._on_log)
            #     print("[WorkerThread] log_signal 连接完成")
            
            app_name = VBAHandlerFactory.get_file_type_name(self.file_type)
            print(f"[WorkerThread] app_name: {app_name}, file_type: {self.file_type}")
            self.log_signal.emit(f"正在初始化{app_name}应用程序...")

            print("[WorkerThread] 调用 handler.initialize()")
            # 初始化处理器
            if not self.handler.initialize():
                self.finished.emit(False, f"{app_name}应用程序初始化失败")
                return

            print("[WorkerThread] initialize 成功，准备打开文档")
            self.log_signal.emit(f"正在打开文件: {self.office_file}")
            
            if self.file_type == FileType.WORD:
                print(f"[WorkerThread] 准备打开Word文档: {self.office_file}")
                if not self.handler.open_document(self.office_file):
                    self.finished.emit(False, f"无法打开{app_name}文档或文档不包含VBA代码")
                    return
                print("[WorkerThread] Word文档打开成功")
            elif self.file_type == FileType.EXCEL:
                if not self.handler.open_workbook(self.office_file):
                    self.finished.emit(False, f"无法打开{app_name}工作簿或工作簿不包含VBA代码")
                    return
            elif self.file_type == FileType.POWERPOINT:
                if not self.handler.open_presentation(self.office_file):
                    self.finished.emit(False, f"无法打开{app_name}演示文稿或演示文稿不包含VBA代码")
                    return

            self.log_signal.emit(f"文档已打开，准备执行 {self.task_type} 任务...")
            
            if self.task_type == 'export':
                self._do_export()
            elif self.task_type == 'import':
                self._do_import()
            elif self.task_type == 'remove':
                self._do_remove()

        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            self.log_signal.emit(f"WorkerThread异常: {str(e)}")
            self.log_signal.emit(f"堆栈: {error_detail}")
            self.finished.emit(False, f"操作失败: {str(e)}\n{error_detail}")
        finally:
            self.log_signal.emit("WorkerThread 进入清理阶段...")
            if self.handler:
                try:
                    if self.file_type == FileType.WORD:
                        self.handler.close_document()
                    elif self.file_type == FileType.EXCEL:
                        self.handler.close_workbook()
                    elif self.file_type == FileType.POWERPOINT:
                        self.handler.close_presentation()
                    self.handler.quit()
                except Exception as e:
                    self.log_signal.emit(f"清理时出错: {e}")
            self.log_signal.emit("WorkerThread 清理完成")

    def _do_export(self):
        """执行导出操作"""
        self.log_signal.emit(f"正在导出 {len(self.components)} 个组件...")
        success = self.handler.export_vba(self.vba_folder, self.components)
        if success:
            self.finished.emit(True, f"成功导出 {len(self.components)} 个VBA组件")
        else:
            self.finished.emit(False, "导出失败")

    def _do_import(self):
        """执行导入操作"""
        self.log_signal.emit(f"正在导入 {len(self.components)} 个组件...")
        success = self.handler.import_vba(self.vba_folder, self.components)
        if success:
            self.finished.emit(True, f"成功导入 {len(self.components)} 个VBA组件")
        else:
            self.finished.emit(False, "导入失败")

    def _do_remove(self):
        """执行清除VBA操作"""
        # 先检查文档是否有VBA代码
        self.log_signal.emit("正在检查VBA代码...")
        components = self.handler.get_vba_components()

        if len(components) > 0:
            # 有VBA代码，执行完整清除
            self.log_signal.emit(f"发现 {len(components)} 个VBA组件，正在清除...")
            success = self.handler.remove_all_vba()
            if success:
                self.finished.emit(True, f"成功清除 {len(components)} 个VBA代码及文档属性")
            else:
                self.finished.emit(False, "清除VBA失败")
        else:
            # 没有VBA代码，仅清除文档属性
            self.log_signal.emit("文档中没有VBA代码，仅清除文档属性...")
            success = self.handler.clear_document_properties_only()
            if success:
                self.finished.emit(True, "成功清除文档属性（无VBA代码）")
            else:
                self.finished.emit(False, "清除文档属性失败")

    def _on_log(self, msg):
        """处理日志消息"""
        print(f"[Handler] {msg}")


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()
        self.office_file = ""
        self.vba_folder = ""
        self.file_type = FileType.WORD  # 默认文件类型
        self.document_components = []  # 文档中的VBA组件
        self.folder_components = []    # 文件夹中的VBA组件
        self.worker_thread = None       # 用于导出/导入/清除操作
        self.refresh_worker = None      # 用于刷新组件列表

        # 初始化日志
        self.logger = None

        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("VBA导入工具 - 支持 Word/Excel/PPT")
        self.setGeometry(100, 100, 850, 700)

        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout()
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(15, 15, 15, 15)
        central_widget.setLayout(main_layout)

        # 添加标题
        title_label = QLabel("VBA导入/导出工具")
        title_font = QFont("Microsoft YaHei", 16, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # 添加分隔线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)

        # 文件类型选择区
        file_type_group = self._create_file_type_section()
        main_layout.addWidget(file_type_group)

        # Office文件选择区
        office_file_group = self._create_office_file_section()
        main_layout.addWidget(office_file_group)

        # VBA文件夹选择区
        folder_group = self._create_folder_section()
        main_layout.addWidget(folder_group)

        # 组件列表区
        components_group = self._create_components_section()
        main_layout.addWidget(components_group)

        # 操作按钮区
        buttons_group = self._create_buttons_section()
        main_layout.addWidget(buttons_group)

        # 日志输出区
        log_group = self._create_log_section()
        main_layout.addWidget(log_group)

        # 状态栏
        self.statusBar().showMessage("就绪")

    def _create_file_type_section(self):
        """创建文件类型选择区域"""
        group = QGroupBox("文件类型")
        layout = QHBoxLayout()

        layout.addWidget(QLabel("选择Office应用:"))

        self.file_type_combo = QComboBox()
        self.file_type_combo.addItem("Word", FileType.WORD)
        self.file_type_combo.addItem("Excel", FileType.EXCEL)
        self.file_type_combo.addItem("PowerPoint", FileType.POWERPOINT)
        self.file_type_combo.currentIndexChanged.connect(self.on_file_type_changed)

        layout.addWidget(self.file_type_combo)
        layout.addStretch()

        # 添加说明标签
        info_label = QLabel("选择要处理的Office文件类型")
        info_label.setStyleSheet("color: gray; font-size: 9pt;")
        layout.addWidget(info_label)

        group.setLayout(layout)
        return group

    def _create_office_file_section(self):
        """创建Office文件选择区域"""
        group = QGroupBox("Office文件")
        layout = QHBoxLayout()

        self.office_file_label = QLabel("未选择文件")
        self.office_file_label.setMinimumWidth(400)
        self.office_file_label.setStyleSheet("color: gray;")

        self.btn_select_file = QPushButton("选择文件")
        self.btn_select_file.clicked.connect(self.select_office_file)

        self.file_type_label = QLabel("Word文件:")
        layout.addWidget(self.file_type_label)
        layout.addWidget(self.office_file_label, 1)
        layout.addWidget(self.btn_select_file)

        group.setLayout(layout)
        return group

    def _create_folder_section(self):
        """创建VBA文件夹选择区域"""
        group = QGroupBox("VBA文件夹")
        layout = QHBoxLayout()

        self.vba_folder_label = QLabel("未选择文件夹")
        self.vba_folder_label.setMinimumWidth(400)
        self.vba_folder_label.setStyleSheet("color: gray;")

        self.btn_select_folder = QPushButton("选择文件夹")
        self.btn_select_folder.clicked.connect(self.select_vba_folder)

        layout.addWidget(QLabel("VBA文件夹:"))
        layout.addWidget(self.vba_folder_label, 1)
        layout.addWidget(self.btn_select_folder)

        group.setLayout(layout)
        return group

    def _create_components_section(self):
        """创建组件列表区域"""
        group = QGroupBox("VBA组件列表")
        layout = QVBoxLayout()

        self.components_list = QListWidget()
        self.components_list.setMinimumHeight(150)
        self.components_list.setSelectionMode(QListWidget.MultiSelection)

        layout.addWidget(self.components_list)

        # 按钮：刷新组件列表
        btn_layout = QHBoxLayout()
        self.btn_refresh = QPushButton("刷新组件列表")
        self.btn_refresh.clicked.connect(self.refresh_components)
        self.btn_refresh.setEnabled(False)

        btn_layout.addWidget(self.btn_refresh)
        btn_layout.addStretch()

        layout.addLayout(btn_layout)

        group.setLayout(layout)
        return group

    def _create_buttons_section(self):
        """创建操作按钮区域"""
        group = QGroupBox("操作")
        layout = QHBoxLayout()

        self.btn_export = QPushButton("导出VBA")
        self.btn_export.clicked.connect(self.export_vba)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setEnabled(False)

        self.btn_import = QPushButton("导入VBA")
        self.btn_import.clicked.connect(self.import_vba)
        self.btn_import.setMinimumHeight(40)
        self.btn_import.setEnabled(False)

        self.btn_remove = QPushButton("清除VBA")
        self.btn_remove.clicked.connect(self.remove_vba)
        self.btn_remove.setMinimumHeight(40)
        self.btn_remove.setEnabled(False)
        self.btn_remove.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
        """)

        layout.addStretch()
        layout.addWidget(self.btn_export)
        layout.addWidget(self.btn_import)
        layout.addWidget(self.btn_remove)

        group.setLayout(layout)
        return group

    def _create_log_section(self):
        """创建日志输出区域"""
        group = QGroupBox("日志输出")
        layout = QVBoxLayout()

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                font-family: Consolas, Courier New;
                font-size: 10pt;
            }
        """)

        # 初始化日志系统
        self.logger = setup_logger("VBA工具", text_widget=self.log_text)
        self.logger.info("程序启动")

        layout.addWidget(self.log_text)

        group.setLayout(layout)
        return group

    def on_file_type_changed(self, index):
        """文件类型改变时的处理"""
        self.file_type = self.file_type_combo.currentData()
        
        # 更新标签
        file_names = {
            FileType.WORD: "Word文件:",
            FileType.EXCEL: "Excel工作簿:",
            FileType.POWERPOINT: "PowerPoint演示文稿:"
        }
        self.file_type_label.setText(file_names[self.file_type])
        
        # 清空已选择的文件
        self.office_file = ""
        self.office_file_label.setText("未选择文件")
        self.office_file_label.setStyleSheet("color: gray;")
        
        # 清空组件列表
        self.components_list.clear()
        self.document_components = []
        self.folder_components = []
        
        self._update_buttons_state()
        self.logger.info(f"已切换到{VBAHandlerFactory.get_file_type_name(self.file_type)}模式")

    def select_office_file(self):
        """选择Office文件"""
        file_filter = VBAHandlerFactory.get_file_filter(self.file_type)
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"选择{VBAHandlerFactory.get_file_type_name(self.file_type)}文件",
            "",
            file_filter
        )

        if file_path:
            self.office_file = file_path
            self.office_file_label.setText(file_path)
            self.office_file_label.setStyleSheet("color: black;")
            self.logger.info(f"已选择文件: {file_path}")
            self.btn_refresh.setEnabled(True)
            self._update_buttons_state()
            self.refresh_components()

    def select_vba_folder(self):
        """选择VBA文件夹"""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "选择VBA文件夹",
            ""
        )

        if folder_path:
            self.vba_folder = folder_path
            self.vba_folder_label.setText(folder_path)
            self.vba_folder_label.setStyleSheet("color: black;")
            self.logger.info(f"已选择VBA文件夹: {folder_path}")
            self._update_buttons_state()
            self.refresh_components()

    def _update_buttons_state(self):
        """更新按钮状态"""
        has_office_file = bool(self.office_file)
        has_vba_folder = bool(self.vba_folder)

        # 导出需要Office文件
        self.btn_export.setEnabled(has_office_file)

        # 导入需要Office文件和VBA文件夹
        self.btn_import.setEnabled(has_office_file and has_vba_folder)

        # 清除VBA只需要Office文件
        self.btn_remove.setEnabled(has_office_file)

    def refresh_components(self):
        """刷新组件列表"""
        # 先清空显示
        self.components_list.clear()
        self.document_components = []
        self.folder_components = []

        # 读取VBA文件夹中的组件（这个可以在主线程完成）
        if self.vba_folder and os.path.exists(self.vba_folder):
            self._load_folder_components()

        # 读取Office文档中的VBA组件（使用后台线程）
        if self.office_file and os.path.exists(self.office_file):
            self._load_document_components_threaded()

    def _load_document_components_threaded(self):
        """使用后台线程加载Office文档中的VBA组件"""
        self._set_buttons_enabled(False)
        self.logger.info("正在读取VBA组件...")

        try:
            self.refresh_worker = RefreshWorkerThread(
                self.office_file,
                self.file_type
            )
            self.refresh_worker.log_signal.connect(self._on_refresh_log)
            self.refresh_worker.finished.connect(self._on_refresh_finished)
            self.refresh_worker.start()
        except Exception as e:
            import traceback
            self.logger.error(f"启动刷新线程失败: {e}")
            self.logger.error(traceback.format_exc())
            self._set_buttons_enabled(True)

    def _on_refresh_log(self, message):
        """处理刷新线程的日志"""
        self.logger.info(message)

    def _on_refresh_finished(self, components, error_msg):
        """处理刷新线程完成"""
        self._set_buttons_enabled(True)

        if error_msg:
            self.logger.error(error_msg)
        else:
            self.document_components = components
            self.logger.info(f"成功读取 {len(components)} 个VBA组件")

        self._display_components()

    def _load_document_components(self):
        """加载Office文档中的VBA组件（已弃用，使用_load_document_components_threaded）"""
        # 此方法已不再使用，所有操作已移至后台线程
        pass

    def _load_folder_components(self):
        """加载VBA文件夹中的组件"""
        try:
            self.logger.info("正在读取VBA文件夹中的组件...")
            # 复用 word_handler 中的 scan_vba_folder 函数
            from core.word_handler import scan_vba_folder
            self.folder_components = scan_vba_folder(self.vba_folder)
            self.logger.info(f"发现 {len(self.folder_components)} 个VBA文件")
        except Exception as e:
            self.logger.error(f"读取文件夹失败: {e}")

    def _display_components(self):
        """显示组件列表"""
        self.components_list.clear()
        
        app_name = VBAHandlerFactory.get_file_type_name(self.file_type)

        # 显示文档中的组件
        if self.document_components:
            header_item = QListWidgetItem(f"=== {app_name}文档中的VBA组件 ===")
            header_item.setFlags(Qt.NoItemFlags)
            header_item.setBackground(Qt.lightGray)
            self.components_list.addItem(header_item)

            for comp in self.document_components:
                item = QListWidgetItem(comp.display_name)
                item.setData(Qt.UserRole, ('document', comp))
                item.setCheckState(Qt.Checked)
                self.components_list.addItem(item)

        # 显示文件夹中的组件
        if self.folder_components:
            header_item = QListWidgetItem("=== VBA文件夹中的组件 ===")
            header_item.setFlags(Qt.NoItemFlags)
            header_item.setBackground(Qt.lightGray)
            self.components_list.addItem(header_item)

            for comp in self.folder_components:
                item = QListWidgetItem(comp.display_name)
                item.setData(Qt.UserRole, ('folder', comp))
                item.setCheckState(Qt.Checked)
                self.components_list.addItem(item)

    def get_selected_components(self):
        """获取选中的组件"""
        selected = {'document': [], 'folder': []}

        for i in range(self.components_list.count()):
            item = self.components_list.item(i)
            if item.checkState() == Qt.Checked:
                data = item.data(Qt.UserRole)
                if data:
                    source, comp = data
                    selected[source].append(comp)

        return selected

    def export_vba(self):
        """导出VBA"""
        selected = self.get_selected_components()
        components_to_export = selected['document']

        if not components_to_export:
            QMessageBox.warning(self, "警告", "没有选择要导出的组件")
            return

        if not self.vba_folder:
            QMessageBox.warning(self, "警告", "请先选择VBA文件夹")
            return

        # 显示确认对话框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("确认导出")
        msg.setText("确认要导出以下VBA组件吗？")

        # 构建组件列表
        component_list = "\n".join([f"• {c.display_name}" for c in components_to_export])
        msg.setInformativeText(
            f"导出组件:\n{component_list}\n\n"
            f"目标文件夹: {self.vba_folder}"
        )
        msg.addButton("确认", QMessageBox.AcceptRole)
        msg.addButton("取消", QMessageBox.RejectRole)

        if msg.exec_() != QMessageBox.AcceptRole:
            return

        # 执行导出
        self._run_task('export', components_to_export)

    def import_vba(self):
        """导入VBA"""
        selected = self.get_selected_components()
        components_to_import = selected['folder']

        if not components_to_import:
            QMessageBox.warning(self, "警告", "没有选择要导入的组件")
            return

        if not self.office_file:
            QMessageBox.warning(self, "警告", "请先选择Office文件")
            return

        # 显示确认对话框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("确认导入")
        msg.setText("确认要导入以下VBA组件吗？")

        # 构建组件列表
        component_list = "\n".join([f"• {c.display_name}" for c in components_to_import])
        app_name = VBAHandlerFactory.get_file_type_name(self.file_type)
        msg.setInformativeText(
            f"导入组件:\n{component_list}\n\n"
            f"目标文档: {self.office_file} ({app_name})"
        )
        msg.addButton("确认", QMessageBox.AcceptRole)
        msg.addButton("取消", QMessageBox.RejectRole)

        if msg.exec_() != QMessageBox.AcceptRole:
            return

        # 执行导入
        self._run_task('import', components_to_import)

    def remove_vba(self):
        """清除VBA"""
        if not self.office_file:
            QMessageBox.warning(self, "警告", "请先选择Office文件")
            return

        # 显示确认对话框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("确认清除VBA")
        msg.setText("确定要清除文档中的所有VBA代码吗？")

        app_name = VBAHandlerFactory.get_file_type_name(self.file_type)
        msg.setInformativeText(
            f"此操作将删除 {app_name} 文档中的所有VBA代码，且无法恢复！\n\n"
            f"文件: {self.office_file}"
        )
        msg.addButton("确认清除", QMessageBox.AcceptRole)
        msg.addButton("取消", QMessageBox.RejectRole)

        if msg.exec_() != QMessageBox.AcceptRole:
            return

        # 执行清除
        self._run_task('remove', [])

    def _run_task(self, task_type, components):
        """执行后台任务"""
        # 检查是否有其他操作正在进行
        if self.worker_thread and self.worker_thread.isRunning():
            QMessageBox.warning(self, "警告", "当前有其他操作正在进行，请等待完成后再执行新操作")
            return

        self._set_buttons_enabled(False)

        self.logger.info(f"创建WorkerThread，task_type={task_type}")
        
        try:
            self.worker_thread = WorkerThread(
                task_type,
                self.office_file,
                self.vba_folder,
                self.file_type,
                components
            )
            self.logger.info("WorkerThread创建成功，连接信号...")
            self.worker_thread.log_signal.connect(self._on_log)
            self.worker_thread.finished.connect(self._on_task_finished)
            self.logger.info("信号连接成功，启动线程...")
            self.worker_thread.start()
            self.logger.info("WorkerThread已启动")
        except Exception as e:
            import traceback
            error_detail = traceback.format_exc()
            self.logger.error(f"创建或启动线程失败: {e}")
            self.logger.error(f"堆栈: {error_detail}")
            self._set_buttons_enabled(True)
            QMessageBox.critical(self, "错误", f"启动任务失败: {e}\n{error_detail}")

        self.logger.info(f"开始执行{task_type}任务...")

    def _on_log(self, message):
        """处理日志消息"""
        self.logger.info(message)

    def _on_task_finished(self, success, message):
        """处理任务完成"""
        self._set_buttons_enabled(True)

        if success:
            self.logger.info(message)
            QMessageBox.information(self, "成功", message)
        else:
            self.logger.error(message)
            QMessageBox.critical(self, "错误", message)

        self.statusBar().showMessage("就绪")

    def _set_buttons_enabled(self, enabled):
        """设置按钮启用状态"""
        self.btn_select_file.setEnabled(enabled)
        self.btn_select_folder.setEnabled(enabled)
        self.btn_refresh.setEnabled(enabled and bool(self.office_file))
        self.btn_export.setEnabled(enabled and bool(self.office_file))
        self.btn_import.setEnabled(enabled and bool(self.office_file) and bool(self.vba_folder))
        self.btn_remove.setEnabled(enabled and bool(self.office_file))
        self.file_type_combo.setEnabled(enabled)

        if not enabled:
            self.statusBar().showMessage("正在处理...")

    def closeEvent(self, event):
        """窗口关闭事件"""
        self.logger.info("程序退出")
        event.accept()


def main():
    """主函数"""
    import sys
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

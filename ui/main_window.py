# -*- coding: utf-8 -*-
"""
PyQt主窗口 - VBA导入工具主界面
"""
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QTextEdit,
    QListWidget, QListWidgetItem, QCheckBox,
    QMessageBox, QFileDialog, QGroupBox,
    QProgressBar, QApplication, QFrame
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QIcon
import os
from core.word_handler import WordVBAHandler, scan_vba_folder
from core.vba_component import VBAComponent
from utils.logger import setup_logger, get_logger


class WorkerThread(QThread):
    """后台工作线程"""
    finished = pyqtSignal(bool, str)
    progress = pyqtSignal(str)
    log_signal = pyqtSignal(str)

    def __init__(self, task_type, word_file, vba_folder, components=None):
        super().__init__()
        self.task_type = task_type  # 'export' or 'import'
        self.word_file = word_file
        self.vba_folder = vba_folder
        self.components = components or []
        self.handler = None

    def run(self):
        try:
            self.handler = WordVBAHandler()
            self.log_signal.emit("正在初始化Word应用程序...")

            if not self.handler.initialize():
                self.finished.emit(False, "Word应用程序初始化失败")
                return

            self.log_signal.emit(f"正在打开文档: {self.word_file}")

            if not self.handler.open_document(self.word_file):
                self.finished.emit(False, "无法打开Word文档或文档不包含VBA代码")
                return

            if self.task_type == 'export':
                self._do_export()
            elif self.task_type == 'import':
                self._do_import()

        except Exception as e:
            self.finished.emit(False, f"操作失败: {str(e)}")
        finally:
            if self.handler:
                self.handler.close_document()
                self.handler.quit()

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


class MainWindow(QMainWindow):
    """主窗口类"""

    def __init__(self):
        super().__init__()
        self.word_file = ""
        self.vba_folder = ""
        self.document_components = []  # Word文档中的VBA组件
        self.folder_components = []    # 文件夹中的VBA组件
        self.worker_thread = None

        # 初始化日志
        self.logger = None

        self.init_ui()

    def init_ui(self):
        """初始化UI"""
        self.setWindowTitle("VBA导入工具")
        self.setGeometry(100, 100, 800, 700)

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

        # Word文件选择区
        word_group = self._create_word_file_section()
        main_layout.addWidget(word_group)

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

    def _create_word_file_section(self):
        """创建Word文件选择区域"""
        group = QGroupBox("Word文件")
        layout = QHBoxLayout()

        self.word_file_label = QLabel("未选择文件")
        self.word_file_label.setMinimumWidth(400)
        self.word_file_label.setStyleSheet("color: gray;")

        self.btn_select_word = QPushButton("选择Word文件")
        self.btn_select_word.clicked.connect(self.select_word_file)

        layout.addWidget(QLabel("Word文件:"))
        layout.addWidget(self.word_file_label, 1)
        layout.addWidget(self.btn_select_word)

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

        layout.addStretch()
        layout.addWidget(self.btn_export)
        layout.addWidget(self.btn_import)

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

    def select_word_file(self):
        """选择Word文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word文件",
            "",
            "Word文件 (*.docm *.doc);;所有文件 (*.*)"
        )

        if file_path:
            self.word_file = file_path
            self.word_file_label.setText(file_path)
            self.word_file_label.setStyleSheet("color: black;")
            self.logger.info(f"已选择Word文件: {file_path}")
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
        has_word_file = bool(self.word_file)
        has_vba_folder = bool(self.vba_folder)

        # 导出需要Word文件
        self.btn_export.setEnabled(has_word_file)

        # 导入需要Word文件和VBA文件夹
        self.btn_import.setEnabled(has_word_file and has_vba_folder)

    def refresh_components(self):
        """刷新组件列表"""
        self.components_list.clear()
        self.document_components = []
        self.folder_components = []

        # 读取Word文档中的VBA组件
        if self.word_file and os.path.exists(self.word_file):
            self._load_word_components()

        # 读取VBA文件夹中的组件
        if self.vba_folder and os.path.exists(self.vba_folder):
            self._load_folder_components()

        self._display_components()

    def _load_word_components(self):
        """加载Word文档中的VBA组件"""
        try:
            self.logger.info("正在读取Word文档中的VBA组件...")
            handler = WordVBAHandler()

            if not handler.initialize():
                self.logger.error("Word应用程序初始化失败")
                return

            if not handler.open_document(self.word_file):
                self.logger.warning("无法打开文档或文档不包含VBA代码")
                handler.quit()
                return

            self.document_components = handler.get_vba_components()
            handler.close_document()
            handler.quit()

            self.logger.info(f"发现 {len(self.document_components)} 个VBA组件")

        except Exception as e:
            self.logger.error(f"读取VBA组件失败: {e}")

    def _load_folder_components(self):
        """加载VBA文件夹中的组件"""
        try:
            self.logger.info("正在读取VBA文件夹中的组件...")
            self.folder_components = scan_vba_folder(self.vba_folder)
            self.logger.info(f"发现 {len(self.folder_components)} 个VBA文件")
        except Exception as e:
            self.logger.error(f"读取文件夹失败: {e}")

    def _display_components(self):
        """显示组件列表"""
        self.components_list.clear()

        # 显示Word文档中的组件
        if self.document_components:
            header_item = QListWidgetItem("=== Word文档中的VBA组件 ===")
            header_item.setFlags(Qt.NoItemFlags)
            header_item.setBackground(Qt.lightGray)
            self.components_list.addItem(header_item)

            for comp in self.document_components:
                item = QListWidgetItem(comp.display_name)
                item.setData(Qt.UserRole, ('word', comp))
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
        selected = {'word': [], 'folder': []}

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
        components_to_export = selected['word']

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

        if not self.word_file:
            QMessageBox.warning(self, "警告", "请先选择Word文件")
            return

        # 显示确认对话框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("确认导入")
        msg.setText("确认要导入以下VBA组件吗？")

        # 构建组件列表
        component_list = "\n".join([f"• {c.display_name}" for c in components_to_import])
        msg.setInformativeText(
            f"导入组件:\n{component_list}\n\n"
            f"目标文档: {self.word_file}"
        )
        msg.addButton("确认", QMessageBox.AcceptRole)
        msg.addButton("取消", QMessageBox.RejectRole)

        if msg.exec_() != QMessageBox.AcceptRole:
            return

        # 执行导入
        self._run_task('import', components_to_import)

    def _run_task(self, task_type, components):
        """执行后台任务"""
        self._set_buttons_enabled(False)

        self.worker_thread = WorkerThread(
            task_type,
            self.word_file,
            self.vba_folder,
            components
        )
        self.worker_thread.log_signal.connect(self._on_log)
        self.worker_thread.finished.connect(self._on_task_finished)
        self.worker_thread.start()

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
        self.btn_select_word.setEnabled(enabled)
        self.btn_select_folder.setEnabled(enabled)
        self.btn_refresh.setEnabled(enabled and bool(self.word_file))
        self.btn_export.setEnabled(enabled and bool(self.word_file))
        self.btn_import.setEnabled(enabled and bool(self.word_file) and bool(self.vba_folder))

        if not enabled:
            self.statusBar().showMessage("正在处理...")

    def closeEvent(self, event):
        """窗口关闭事件"""
        # 确保关闭Word进程
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

# -*- coding: utf-8 -*-
"""
VBA 代码测试脚本
功能：自动测试学生作业保护系统的VBA代码
用法：python test_vba.py
"""

import os
import sys
import time
import shutil
from pathlib import Path

# 检查并安装依赖
try:
    import win32com.client
    import pythoncom
except ImportError:
    print("正在安装 pywin32...")
    os.system("pip install pywin32")
    import win32com.client
    import pythoncom


class VBAWordTester:
    """Word VBA 代码测试类"""
    
    def __init__(self, template_path: str):
        self.template_path = Path(template_path)
        self.test_doc_path = None
        self.word_app = None
        self.doc = None
        
    def setup(self):
        """准备测试环境"""
        print("=" * 50)
        print("开始测试 VBA 代码")
        print("=" * 50)
        
        # 复制模板文件到测试文件
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        self.test_doc_path = self.template_path.parent / f"测试文档_{timestamp}.docm"
        shutil.copy2(self.template_path, self.test_doc_path)
        print(f"\n[1] 创建测试文档: {self.test_doc_path}")
        
        # 初始化 Word 应用
        print("\n[2] 启动 Word 应用程序...")
        self.word_app = win32com.client.Dispatch("Word.Application")
        self.word_app.Visible = True  # 显示Word窗口，方便调试
        self.word_app.DisplayAlerts = False  # 禁用弹出警告
        
        return True
        
    def test_document_open(self):
        """测试文档打开事件"""
        print("\n[3] 打开测试文档 (触发 Document_Open)...")
        
        try:
            # 打开文档，这会触发 Document_Open 事件
            # 注意：由于 VBA 中有 InputBox，Word会暂停等待输入
            self.doc = self.word_app.Documents.Open(str(self.test_doc_path))
            print("    文档已打开")
            print("    注意: 由于 VBA 中有 InputBox 对话框，请手动输入测试数据")
            print("    学号: TEST001")
            print("    姓名: 测试学生")
            return True
        except Exception as e:
            print(f"    错误: {e}")
            return False
            
    def test_manual_login(self):
        """手动登录测试 - 用户在Word中输入"""
        print("\n[4] 等待用户完成登录...")
        print("    请在弹出的 Word 窗口中：")
        print("    1. 输入学号: TEST001")
        print("    2. 输入姓名: 测试学生")
        print("    3. 点击确定完成登录")
        input("    按回车键继续...")
        
    def test_check_watermark(self):
        """检查水印是否添加成功"""
        print("\n[5] 检查水印...")
        
        try:
            # 检查页眉水印
            header = self.doc.Sections(1).Headers(1)
            print(f"    页眉形状数量: {header.Shapes.Count}")
            
            # 检查页面水印
            print(f"    文档形状数量: {self.doc.Shapes.Count}")
            
            # 列出所有形状名称
            for i, shp in enumerate(self.doc.Shapes):
                print(f"    形状 {i+1}: {shp.Name}")
                
            return True
        except Exception as e:
            print(f"    错误: {e}")
            return False
            
    def test_teacher_reset(self):
        """测试老师重置功能"""
        print("\n[6] 测试老师重置功能...")
        
        try:
            # 调用老师重置宏
            self.doc.VBProject.VBComponents.Import(str(self.template_path))
            # 注意：实际调用需要通过 Run 方法
            # self.word_app.Run("TeacherReset")
            print("    老师重置功能测试需要手动完成")
            return True
        except Exception as e:
            print(f"    注意: {e}")
            print("    (某些功能需要启用宏安全性)")
            return True
            
    def cleanup(self):
        """清理测试环境"""
        print("\n[7] 清理测试环境...")
        
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
                print("    文档已关闭")
        except:
            pass
            
        try:
            if self.word_app:
                self.word_app.Quit()
                print("    Word 已关闭")
        except:
            pass
            
        # 删除测试文件
        try:
            if self.test_doc_path and self.test_doc_path.exists():
                # 等待文件解锁
                time.sleep(1)
                self.test_doc_path.unlink()
                print("    测试文件已删除")
        except:
            print("    注意: 测试文件可能需要手动删除")
            
        print("\n" + "=" * 50)
        print("测试完成!")
        print("=" * 50)
        
    def run_full_test(self):
        """运行完整测试流程"""
        try:
            self.setup()
            self.test_document_open()
            self.test_manual_login()
            self.test_check_watermark()
            # self.test_teacher_reset()  # 可选
            return True
        except Exception as e:
            print(f"\n测试出错: {e}")
            import traceback
            traceback.print_exc()
            return False
        finally:
            self.cleanup()


def check_word_available():
    """检查 Word 是否可用"""
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except:
        return False


def main():
    """主函数"""
    print("VBA 代码自动化测试工具")
    print("-" * 40)
    
    # 检查 Word 是否可用
    if not check_word_available():
        print("错误: 无法启动 Word 应用程序")
        print("请确保已安装 Microsoft Word")
        sys.exit(1)
        
    # 设置测试文件路径
    script_dir = Path(__file__).parent
    template_file = script_dir / "demo02" / "vbacode002" / "学生作业模板.docm"
    
    if not template_file.exists():
        # 尝试查找其他可能的模板文件
        for doc_file in script_dir.rglob("*.docm"):
            template_file = doc_file
            break
            
    if not template_file.exists():
        print(f"错误: 找不到模板文件")
        print(f"请将模板文件放到: {script_dir}")
        sys.exit(1)
        
    print(f"使用模板文件: {template_file}")
    
    # 创建测试器并运行测试
    tester = VBAWordTester(str(template_file))
    tester.run_full_test()


if __name__ == "__main__":
    main()

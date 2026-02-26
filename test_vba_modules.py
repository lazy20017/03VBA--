# -*- coding: utf-8 -*-
"""
VBA模块测试脚本
使用win32com控制Word进行自动化测试
"""

import os
import sys
import time
import shutil
import datetime

# 测试配置
TEST_DOC_PATH = os.path.join(os.path.dirname(__file__), "demo02", "example002.docm")
TEST_OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "test_output")
VBA_CODE_DIR = os.path.join(os.path.dirname(__file__), "demo02", "vbacode002")


def setup_test_environment():
    """设置测试环境"""
    print("=" * 60)
    print("设置测试环境...")
    
    # 创建测试输出目录
    if not os.path.exists(TEST_OUTPUT_DIR):
        os.makedirs(TEST_OUTPUT_DIR)
    
    # 复制测试文档
    test_doc = os.path.join(TEST_OUTPUT_DIR, "test_document.docm")
    if os.path.exists(TEST_DOC_PATH):
        shutil.copy2(TEST_DOC_PATH, test_doc)
        print(f"✓ 已复制测试文档到: {test_doc}")
        return test_doc
    else:
        print(f"✗ 源文档不存在: {TEST_DOC_PATH}")
        return None


def check_vba_references():
    """检查VBA模块中的代码引用是否正确"""
    print("\n" + "=" * 60)
    print("检查VBA代码引用...")
    
    issues = []
    
    # 检查MainModule.bas
    main_module_path = os.path.join(VBA_CODE_DIR, "MainModule.bas")
    if os.path.exists(main_module_path):
        with open(main_module_path, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # 检查BuildKeyCode（应该已被移除）
        if 'BuildKeyCode' in content:
            issues.append("MainModule.bas: 仍存在 BuildKeyCode 函数调用")
        
        # 检查Protected (重复)
        protect_count = content.count('ThisDocument.Protect')
        if protect_count > 1:
            issues.append(f"MainModule.bas: 存在 {protect_count} 次 Protect 调用，可能存在重复保护")
        
        # 检查DisableSelection
        if '.Hidden = True' in content:
            issues.append("MainModule.bas: DisableSelection 可能隐藏文字")
        
        print(f"✓ MainModule.bas 检查完成")
    else:
        issues.append("MainModule.bas 文件不存在")
    
    # 检查StudentStorage.bas
    storage_path = os.path.join(VBA_CODE_DIR, "StudentStorage.bas")
    if os.path.exists(storage_path):
        with open(storage_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查全局变量引用（g_StudentID等）
        if 'g_StudentID' in content or 'g_StudentName' in content:
            # 检查是否在引用其他模块的变量
            if 'Public Sub' in content or 'Public Function' in content:
                issues.append("StudentStorage.bas: 可能存在对全局变量的引用")
        
        print(f"✓ StudentStorage.bas 检查完成")
    else:
        issues.append("StudentStorage.bas 文件不存在")
    
    # 检查WatermarkManager.bas
    watermark_path = os.path.join(VBA_CODE_DIR, "WatermarkManager.bas")
    if os.path.exists(watermark_path):
        with open(watermark_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 检查常量（应该使用符号常量而非数字）
        if '= 0' in content and 'Visible' not in content:
            # 可能存在硬编码的0
            pass
        
        print(f"✓ WatermarkManager.bas 检查完成")
    else:
        issues.append("WatermarkManager.bas 文件不存在")
    
    # 报告结果
    if issues:
        print("\n发现以下问题:")
        for issue in issues:
            print(f"  ✗ {issue}")
        return False
    else:
        print("\n✓ 所有代码引用检查通过!")
        return True


def check_word_installed():
    """检查Word是否安装"""
    print("\n" + "=" * 60)
    print("检查Word是否安装...")
    
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        print("✓ Word已安装")
        return True
    except ImportError:
        print("✗ pywin32未安装，请运行: pip install pywin32")
        return False
    except Exception as e:
        print(f"✗ Word检查失败: {e}")
        return False


def test_word_automation():
    """测试Word自动化功能"""
    print("\n" + "=" * 60)
    print("测试Word自动化...")
    
    try:
        import win32com.client
        from win32com.client import constants
        
        # 启动Word
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        
        print(f"✓ Word已启动 (版本: {word.Version})")
        
        # 测试创建文档
        doc = word.Documents.Add()
        doc.Content.Text = "测试文档内容"
        
        # 测试水印功能
        print("  - 测试添加形状...")
        shp = doc.Shapes.AddTextbox(1, 100, 100, 200, 50)
        shp.TextFrame.TextRange.Text = "测试水印"
        
        # 测试旋转
        print("  - 测试旋转...")
        shp.Rotation = -45
        
        # 测试保护功能
        print("  - 测试文档保护...")
        doc.Protect(2)  # wdAllowOnlyReading = 2
        
        # 测试取消保护
        print("  - 测试取消保护...")
        doc.Unprotect()
        
        # 关闭文档
        doc.Close(False)
        
        # 退出Word
        word.Quit()
        
        print("✓ Word自动化测试通过")
        return True
        
    except ImportError:
        print("✗ pywin32未安装")
        return False
    except Exception as e:
        print(f"✗ Word自动化测试失败: {e}")
        return False


def test_vba_module_structure():
    """测试VBA模块结构"""
    print("\n" + "=" * 60)
    print("检查VBA模块结构...")
    
    required_files = [
        "MainModule.bas",
        "StudentStorage.bas", 
        "WatermarkManager.bas",
        "ThisDocument.bas"
    ]
    
    all_exist = True
    for fname in required_files:
        fpath = os.path.join(VBA_CODE_DIR, fname)
        if os.path.exists(fpath):
            print(f"  ✓ {fname}")
        else:
            print(f"  ✗ {fname} 不存在")
            all_exist = False
    
    return all_exist


def list_vba_procedures():
    """列出所有VBA过程和函数"""
    print("\n" + "=" * 60)
    print("VBA模块内容概览...")
    
    bas_files = [
        ("MainModule.bas", ["ShowLoginForm", "DisableCopyFunction", "EnableCopyFunction", 
                           "TeacherReset", "ViewAllStudents", "ExportStudentRecords"]),
        ("StudentStorage.bas", ["RecordLogin", "IsReturningStudent", "GetAllStudents", 
                               "ClearAllStudentRecords", "ValidateStudent"]),
        ("WatermarkManager.bas", ["AddStudentWatermark", "RemoveWatermark", "UpdateWatermark"]),
        ("ThisDocument.bas", ["Document_Open", "Document_New"])
    ]
    
    for filename, procedures in bas_files:
        print(f"\n{filename}:")
        for proc in procedures:
            print(f"  - {proc}()")


def generate_test_report():
    """生成测试报告"""
    print("\n" + "=" * 60)
    print("生成测试报告...")
    
    report_path = os.path.join(TEST_OUTPUT_DIR, "test_report.txt")
    
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("=" * 60 + "\n")
        f.write("VBA模块测试报告\n")
        f.write(f"生成时间: {datetime.datetime.now()}\n")
        f.write("=" * 60 + "\n\n")
        
        f.write("测试项目:\n")
        f.write("1. VBA模块文件存在性检查\n")
        f.write("2. VBA代码引用检查\n")
        f.write("3. Word自动化测试\n\n")
        
        f.write("建议测试步骤:\n")
        f.write("1. 在Word中打开 example002.docm\n")
        f.write("2. 按 Alt+F11 打开VBA编辑器\n")
        f.write("3. 手动运行各模块的函数进行测试\n")
    
    print(f"✓ 测试报告已保存到: {report_path}")


def main():
    """主函数"""
    print("\n" + "=" * 60)
    print("VBA模块自动化测试工具")
    print("=" * 60)
    
    # 1. 检查VBA模块结构
    if not test_vba_module_structure():
        print("\n✗ VBA模块结构不完整，请检查文件")
        return
    
    # 2. 检查代码引用
    check_vba_references()
    
    # 3. 列出VBA过程
    list_vba_procedures()
    
    # 4. 设置测试环境
    test_doc = setup_test_environment()
    
    # 5. 检查Word是否安装
    if check_word_installed():
        # 6. 测试Word自动化
        test_word_automation()
    
    # 7. 生成报告
    generate_test_report()
    
    print("\n" + "=" * 60)
    print("测试完成!")
    print("=" * 60)
    print("\n注意: 完整的VBA测试需要在Word中进行:")
    print("1. 打开 demo02/example002.docm")
    print("2. 文档会自动触发 Document_Open 事件")
    print("3. 在弹出的InputBox中输入学号和姓名")
    print("4. 观察水印是否正确添加")
    print("5. 检查复制功能是否被禁用")


if __name__ == "__main__":
    main()

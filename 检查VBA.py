# -*- coding: utf-8 -*-
"""
快速检查Word文档中的VBA代码
"""
import os
import win32com.client
import pythoncom

def check_vba_in_document(file_path: str):
    """检查文档中的VBA代码"""
    try:
        pythoncom.CoInitialize()
        
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = True  # 让Word可见，方便对比
        word_app.DisplayAlerts = False
        
        print(f"正在打开: {file_path}")
        doc = word_app.Documents.Open(os.path.abspath(file_path))
        
        print("\n" + "="*50)
        print("VBA项目内容:")
        print("="*50)
        
        vba_project = doc.VBProject
        if vba_project:
            components = vba_project.VBComponents
            print(f"共发现 {len(components)} 个组件:\n")
            
            for i, comp in enumerate(components, 1):
                print(f"--- 组件 {i}: {comp.Name} ---")
                print(f"类型: ", end="")
                
                # 显示类型
                if comp.Type == 1:
                    print("标准模块 (Standard Module)")
                elif comp.Type == 2:
                    print("类模块 (Class Module)")
                elif comp.Type == 3:
                    print("用户窗体 (UserForm)")
                elif comp.Type == 100:
                    print("文档模块 (Document Module)")
                else:
                    print(f"未知类型 ({comp.Type})")
                
                # 显示代码
                try:
                    code_module = comp.CodeModule
                    if code_module and code_module.CountOfLines > 0:
                        code = code_module.Lines(1, code_module.CountOfLines)
                        print("代码内容:")
                        print("-" * 40)
                        print(code[:500] if len(code) > 500 else code)
                        if len(code) > 500:
                            print(f"... (共 {code_module.CountOfLines} 行)")
                    else:
                        print("  (空模块)")
                except Exception as e:
                    print(f"  无法读取代码: {e}")
                print()
        else:
            print("文档没有VBA项目!")
        
        input("\n按回车键关闭Word...")
        
        doc.Close(False)
        word_app.Quit()
        pythoncom.CoUninitialize()
        
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    doc_path = r"E:\pydemo\03VBA工程\demo\example001.docm"
    if os.path.exists(doc_path):
        check_vba_in_document(doc_path)
    else:
        print(f"文件不存在: {doc_path}")

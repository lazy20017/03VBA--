# VBA导入工具

基于PyQt5开发的Word VBA代码导入/导出工具。

## 功能特点

- **导出VBA**: 将Word文档中的VBA代码导出为多个独立文件
- **导入VBA**: 将文件夹中的VBA代码导入到Word文档
- **组件管理**: 支持标准模块(.bas)、类模块(.cls)、窗体(.frm)等类型
- **弹窗确认**: 导入导出前显示确认对话框
- **日志输出**: 实时显示操作日志

## 项目结构

```
VBA工程/
├── vba_import_tool.py          # Main program entry
├── 开发设计文档.md          # 开发设计文档
├── requirements.txt        # Python依赖
├── README.md              # 说明文档
├── ui/
│   ├── __init__.py
│   └── main_window.py     # 主窗口UI
├── core/
│   ├── __init__.py
│   ├── vba_component.py   # VBA组件类
│   └── word_handler.py    # Word VBA处理
└── utils/
    ├── __init__.py
    └── logger.py          # 日志工具
```

## 环境要求

- Windows操作系统
- Microsoft Word 2016 (需要安装)
- Python 3.8+

## 安装步骤

1. 安装Python依赖:

```bash
pip install -r requirements.txt
```

2. 确保已安装Microsoft Word 2016

## 使用方法

1. 运行程序:

```bash
python vba_import_tool.py
```

2. **导出VBA**:
   - 点击"选择Word文件"按钮，选择.docm或.doc文件
   - 点击"选择文件夹"按钮，选择保存VBA文件的文件夹
   - 在组件列表中勾选要导出的组件
   - 点击"导出VBA"按钮
   - 在弹窗确认中点击"确认"

3. **导入VBA**:
   - 点击"选择Word文件"按钮，选择.docm文件
   - 点击"选择文件夹"按钮，选择包含VBA代码的文件夹
   - 在组件列表中勾选要导入的组件
   - 点击"导入VBA"按钮
   - 在弹窗确认中点击"确认"

## VBA文件类型说明

| 类型 | 扩展名 | 说明 |
|------|--------|------|
| 标准模块 | .bas | 常规VBA代码模块 |
| 类模块 | .cls | VBA类定义 |
| 窗体 | .frm | UserForm窗体 |

## 注意事项

- 导出的Word文件需要包含VBA代码（.docm格式）
- 导入时Word文件必须启用宏
- 操作过程中请勿关闭Word程序
- 建议在操作前备份原始文件

## 常见问题

1. **无法打开Word文档**: 确保Word已正确安装，且文档不是只读模式
2. **找不到VBA组件**: 确保Word文档包含VBA代码（.docm格式）
3. **导入失败**: 确保VBA文件夹中的文件格式正确（.bas/.cls/.frm）

' -*- coding: utf-8 -*-
' =====================================================
' ThisDocument 模块 - 文档事件处理
' 功能：处理文档打开、关闭等事件
' =====================================================

Option Explicit

' =====================================================
' 文档事件 - 打开文档时自动运行
' =====================================================

' 文档打开时自动运行
Private Sub Document_Open()
    Call ShowLoginForm
End Sub

' 文档初始化（兼容旧版本Word）
Private Sub Document_New()
    Call ShowLoginForm
End Sub

' 关闭文档时保存学生信息
Private Sub Document_Close()
    ' （学生信息已通过StudentStorage模块保存在书签中）
End Sub

' 文档内容变化时检查（防止复制）
Private Sub Document_ContentControlOnExit(ByVal ContentControl As ContentControl, Cancel As Boolean)
    ' 可以添加内容检查逻辑
End Sub

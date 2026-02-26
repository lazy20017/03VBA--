' -*- coding: utf-8 -*-
' =====================================================
' 水印管理模块
' 功能：为文档添加学生学号和姓名水印
' =====================================================

Option Explicit

' Word/Office 常量声明
Const wdStatisticPages As Long = 2
Const wdAlignParagraphCenter As Long = 1
Const wdAlignParagraphRight As Long = 2
Const wdWrapBehind As Long = 6
Const wdHeaderFooterPrimary As Long = -1
Const msoTextOrientationHorizontal As Long = 1
Const msoFalse As Long = 0
Const wdPrintView As Long = 3

' 水印形状名称
Private Const WATERMARK_SHAPE_NAME As String = "StudentIDWatermark"

' 调试用 - 记录错误
Private Function GetErrInfo() As String
    GetErrInfo = "错误号: " & Err.Number & "  描述: " & Err.Description
End Function

' 添加学生水印
Public Sub AddStudentWatermark(studentID As String, studentName As String)
    Dim doc As Document
    
    Set doc = ThisDocument
    
    ' 先解除文档保护（如果已保护）
    On Error Resume Next
    doc.Unprotect
    If Err.Number <> 0 And Err.Number <> 5380 Then  ' 5380 = 文档未保护
        ' 忽略其他错误继续执行
    End If
    Err.Clear
    On Error GoTo 0
    
    ' 确保文档处于页面视图
    If doc.ActiveWindow.View.Type <> wdPrintView Then
        doc.ActiveWindow.View.Type = wdPrintView
    End If
    
    ' 先删除已有的水印
    Call RemoveWatermark
    
    ' 添加文字水印（对角线倾斜）
    Call AddTextWatermark(studentID, studentName)
    
    ' 设置文件属性（学号和姓名）
    Call SetDocumentProperty(studentID, studentName)
End Sub

' 添加文字水印（对角线倾斜）
Private Sub AddTextWatermark(studentID As String, studentName As String)
    Dim shp As Shape
    Dim watermarkText As String
    Dim doc As Document
    Dim i As Integer
    Dim pageCount As Long
    
    Set doc = ThisDocument
    watermarkText = studentName & "  " & studentID
    
    ' 获取页数
    On Error Resume Next
    pageCount = doc.ComputeStatistics(wdStatisticPages)
    If pageCount = 0 Then pageCount = 1  ' 防止空文档
    
    ' 为每一页添加水印
    For i = 1 To pageCount
        ' 在每页中心添加水印形状
        Set shp = doc.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            Left:=0, Top:=0, Width:=0, Height:=0, _
            Anchor:=doc.Range(0))
        
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextPage
        End If
        
        With shp
            .Name = WATERMARK_SHAPE_NAME & "_" & i
            .TextFrame.TextRange.Text = watermarkText
            .TextFrame.TextRange.Font.Size = 44
            .TextFrame.TextRange.Font.Color = RGB(200, 200, 200)
            .TextFrame.TextRange.Font.Name = "黑体"
            .TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphCenter
            
            ' 设置旋转角度（-45度）
            .Rotation = -45
            
            ' 设置位置和大小（页面居中）
            .Left = 100
            .Top = 200
            .Width = 400
            .Height = 100
            
            ' 设置半透明
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            
            ' 设置版式为文字下方
            .WrapFormat.Type = wdWrapBehind
        End With
        
        ' 调整位置到页面中心
        Call PositionWatermarkOnPage(shp, i)
NextPage:
    Next i
    On Error GoTo 0
End Sub

' 将水印定位到页面中心
Private Sub PositionWatermarkOnPage(shp As Shape, pageNum As Integer)
    Dim doc As Document
    Dim pageWidth As Single, pageHeight As Single
    Dim marginLeft As Single, marginTop As Single
    
    Set doc = ThisDocument
    
    ' 获取页面尺寸（磅）
    pageWidth = doc.PageSetup.PageWidth
    pageHeight = doc.PageSetup.PageHeight
    
    ' 获取页边距
    marginLeft = doc.PageSetup.LeftMargin
    marginTop = doc.PageSetup.TopMargin
    
    ' 计算页面可用的中心和尺寸
    Dim usableWidth As Single
    Dim usableHeight As Single
    usableWidth = pageWidth - marginLeft - doc.PageSetup.RightMargin
    usableHeight = pageHeight - marginTop - doc.PageSetup.BottomMargin
    
    ' 居中定位
    With shp
        .Left = marginLeft + (usableWidth - .Width) / 2
        .Top = marginTop + (usableHeight - .Height) / 2
    End With
End Sub

' 删除水印
Public Sub RemoveWatermark()
    Dim doc As Document
    Dim shp As Shape
    Dim toDelete As Collection
    Dim s As Variant
    
    Set doc = ThisDocument
    Set toDelete = New Collection
    
    ' 收集要删除的形状
    On Error Resume Next
    For Each shp In doc.Shapes
        If shp.Name Like WATERMARK_SHAPE_NAME & "*" Then
            toDelete.Add shp.Name
        End If
    Next shp
    
    ' 删除形状
    For Each s In toDelete
        doc.Shapes(CStr(s)).Delete
    Next s
    
    On Error GoTo 0
End Sub

' 设置文档属性（学号和姓名）- 防止抄袭
Public Sub SetDocumentProperty(studentID As String, studentName As String)
    Dim doc As Document
    
    Set doc = ThisDocument
    
    On Error Resume Next
    
    ' 设置内置属性 - 主题（可用于标识）
    doc.BuiltInDocumentProperties("Subject").Value = studentID & " - " & studentName
    
    ' 设置自定义属性 - 学号
    Call SetCustomProperty(doc, "学生学号", studentID)
    
    ' 设置自定义属性 - 姓名
    Call SetCustomProperty(doc, "学生姓名", studentName)
    
    On Error GoTo 0
End Sub

' 设置自定义属性
Private Sub SetCustomProperty(doc As Document, propName As String, propValue As String)
    Dim props As Object
    Dim i As Integer
    Dim found As Boolean
    
    On Error Resume Next
    
    ' 尝试获取自定义属性集合
    Set props = doc.CustomDocumentProperties
    
    If props Is Nothing Then
        Exit Sub
    End If
    
    ' 检查是否已存在该属性
    found = False
    For i = 1 To props.Count
        If props(i).Name = propName Then
            ' 更新属性值
            props(propName).Value = propValue
            found = True
            Exit For
        End If
    Next i
    
    ' 如果不存在，则添加新属性
    If Not found Then
        ' 添加自定义属性（msoPropertyTypeString = 2）
        doc.CustomDocumentProperties.Add Name:=propName, _
            LinkToContent:=False, _
            Type:=2, _
            Value:=propValue
    End If
    
    On Error GoTo 0
End Sub

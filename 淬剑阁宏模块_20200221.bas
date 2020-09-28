Attribute VB_Name = "淬剑阁宏模块_20200221"
Option Explicit
' 代码说明：
'     【淬剑阁宏模块】系列代码由【淬剑阁】（https://www.cuijiange.com）出品，旨在提高工作效率，减少重复劳动
'     代码授权，本代码参考了部分网友的功能核心代码实现，在其基础上增加了部分功能和错误处理，使用者可以不受限制的进行拷贝、修改以及商业使用
' 作者：淬剑阁阁主-西陵剑魂
' 使用方法：
'     打开Excel，【新建】空白工作簿，按【Alt+F11】打开【Visual Basic for Application编辑器】
'     打开【文件】，选择【导入文件】，浏览到本代码所在位置，选中本文件，点击【确定】导入
'     关闭【Visual Basic for Application编辑器】，点击【另存为】，更改文件类型为【Excel加载宏|*.xlam】，点击确定

' 模块名称：淬剑阁宏模块_20200221
' 模块功能：
' 1、自动调整行高：

' 2、自动调整图片大小：

' 3、按指定字符和间隔分割文本(splitIn)

Sub 自动调整行高()
'     快捷方式：无
'     功能描述：
'         自动调整选中的单元格行高，支持合并单元格，支持同时选择多个单元格区域
'     使用要领：
'         选中一个或多个需要调整行高的单元格，执行宏
    Dim g, addHeightRow, Rowheight_25
    Dim rh As Single, mw As Single
    Dim rng As Range, rrng As Range, n1%, n2%
    Dim aw As Single, rh1 As Single
    Dim m$, n$, k
    Dim ir1, ir2, ic1, ic2
    Dim mySheet As Worksheet
    Dim selectedA As Range
    Dim wrkSheet As Worksheet
    
    Application.ScreenUpdating = False
    Set mySheet = ActiveSheet
    On Error Resume Next
    Err.Number = 0
    Set selectedA = Application.Intersect(ActiveWindow.RangeSelection, mySheet.UsedRange)
    selectedA.Activate
    If Err.Number <> 0 Then
        g = MsgBox("请先选择需要'最合适行高'的行!", vbInformation)
        Return
    End If
    selectedA.EntireRow.AutoFit
    Set wrkSheet = ActiveWorkbook.Worksheets.Add
    For Each rrng In selectedA
        If rrng.Address <> rrng.MergeArea.Address Then
            If rrng.Address = rrng.MergeArea.Item(1).Address Then
                
                'If (Application.Intersect(selectedA, rrng).Address <> rrng.Address) Then
                '    GoTo gotoNext
                'End If
                
                Dim tempCell As Range
                Dim width As Double
                Dim tempcol
                width = 0
                For Each tempcol In rrng.MergeArea.Columns
                    width = width + tempcol.ColumnWidth
                Next
                wrkSheet.Columns(1).WrapText = True
                wrkSheet.Columns(1).ColumnWidth = width
                wrkSheet.Columns(1).Font.Name = rrng.Font.Name
                wrkSheet.Columns(1).Font.Size = rrng.Font.Size
                wrkSheet.Cells(1, 1).Value = rrng.Value
                wrkSheet.Activate
                wrkSheet.Cells(1, 1).RowHeight = 0
                wrkSheet.Cells(1, 1).EntireRow.Activate
                wrkSheet.Cells(1, 1).EntireRow.AutoFit
                mySheet.Activate
                rrng.Activate
                If (rrng.RowHeight < wrkSheet.Cells(1, 1).RowHeight) Then
                    Dim tempHeight As Double
                    Dim tempCount As Integer
                    tempHeight = wrkSheet.Cells(1, 1).RowHeight
                    tempCount = rrng.MergeArea.Rows.Count
                    For Each addHeightRow In rrng.MergeArea.Rows
                        Dim rng As Range
                        If (addHeightRow.RowHeight < tempHeight / tempCount) Then
                            addHeightRow.RowHeight = tempHeight / tempCount
                        End If
                        tempHeight = tempHeight - addHeightRow.RowHeight
                        tempCount = tempCount - 1
                    Next
                End If
            End If
        End If
    Next
    Application.DisplayAlerts = False            '删除工作表警告提示去消
    wrkSheet.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Public Sub 自动调整图片大小()
Attribute 自动调整图片大小.VB_ProcData.VB_Invoke_Func = "Q\n14"

'     快捷方式：Ctrl+Shift+Q，如不需要快捷方式，将下方 Attribute 自动调整图片大小.VB_ProcData.VB_Invoke_Func = "Q\n14" 删除
'     功能描述：
'         自动调整激活工作表中选中的图形大小（如未选中任何图形，则默认调整所有图形），
'         使其适应图形左上角所对应的单元格的宽高，调整时保持原宽高比不变
'     使用要领：
'         选中单个或多个需要调整大小的图片或图形，执行宏，如不选则针对激活工作表所有图形进行调整
'Attribute 自动调整图片大小.VB_ProcData.VB_Invoke_Func = "Q\n14"

    If TypeName(Selection) = "Range" Then
        ActiveSheet.Shapes.SelectAll
        If TypeName(Selection) = "Range" Then
            MsgBox "本工作表无可调整对象"
            Exit Sub
        End If
    End If
ChangePic Selection
End Sub
Private Sub ChangePic(Optional obj)
'
' ChangePic 宏

    Dim s, typ

    If obj Is Nothing Then
        Set s = Selection
        Else
        Set s = obj
        End If
    Dim scaleValue
    scaleValue = 0.95   '1＞scaleValue＞0
    On Error Resume Next
        If s.TopLeftCell = 0 Then
        End If                  '判断是否是图形
        If Err.Number = 0 Then
            typ = "Shape"
        Else
            Err.Clear
            If s.Count <> 0 Then
            End If              '判断是否是集合
            If Err.Number = 0 Then
                typ = "Arrow"
            Else
                MsgBox "请选择【单个】或【多个】图片或形状"
                Exit Sub
            End If
            
        End If
        Err.Clear
    On Error GoTo 0
    If typ = "Shape" Then
        Dim PicHpW As Double
        Dim RngHpW As Double
        Dim rng As Range
        With s  '可用
            
            If .TopLeftCell.MergeCells = True Then
                Set rng = .TopLeftCell.MergeArea
            Else
                Set rng = .TopLeftCell
            End If
        RngHpW = rng.Height / rng.width
        PicHpW = .Height / .width
        .ShapeRange.LockAspectRatio = msoFalse
        .Placement = xlMove
        If RngHpW >= PicHpW Then
            .ShapeRange.width = scaleValue * rng.width
            .ShapeRange.Height = .ShapeRange.width * PicHpW
        Else
            .ShapeRange.Height = scaleValue * rng.Height
            .ShapeRange.width = .ShapeRange.Height / PicHpW
        End If
        .ShapeRange.Left = rng.width / 2 - .ShapeRange.width / 2 + rng.Left
        .ShapeRange.Top = rng.Height / 2 - .ShapeRange.Height / 2 + rng.Top
        End With
    ElseIf typ = "Arrow" Then
        Dim sv
        For Each sv In s
                ChangePic sv
        Next
    End If
End Sub

'splitIn(ByVal str As String, Optional splitCharsNumber = 4, Optional splitChar = vbCrLf)
'使用指定分隔符和指定间隔数对字符串进行分割，返回分割后的字符串
'示例：
    'splitIn("1234567890",4," ") 返回值："1234 5678 90"
    '
'str                必选    需要进行分割的字符串
'SplitCharsNumber   可选    需要分割的间隔数    默认值  4
'SplitChar          可选    分割时候使用的字符  默认值  VbCrLf（回车换行符）
Function splitIn(ByVal str As String, Optional splitCharsNumber As Integer = 4, Optional splitChar As String = vbCrLf) As String
    Dim RowNumber As Integer, StrLen As Integer
    StrLen = Len(str)
    If splitCharsNumber <= 0 Then
        Err.Raise vbObjectError + 513, , "分割间隔不能为0或负值"
    End If
    If StrLen > 0 Then
        For RowNumber = Int((StrLen - 1) / splitCharsNumber) * splitCharsNumber + 1 To 1 Step -splitCharsNumber
            splitIn = Mid(str, RowNumber, 4) & splitChar & splitIn
        Next
        splitIn = Left(splitIn, Len(splitIn) - Len(splitChar))
    End If
End Function
Private Sub splitIn_test()
    Dim arr(), RowNumber, ColumnNumber, t
    't = Timer
    arr = UsedRange
    For RowNumber = LBound(arr) To UBound(arr)
        For ColumnNumber = LBound(arr, 2) To UBound(arr, 2)
            arr(RowNumber, ColumnNumber) = splitIn(arr(RowNumber, ColumnNumber))
        Next
    Next
    UsedRange = arr
    'Debug.Print Timer - t
End Sub

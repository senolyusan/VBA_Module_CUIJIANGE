Attribute VB_Name = "�㽣���ģ��_20200221"
Option Explicit
' ����˵����
'     ���㽣���ģ�顿ϵ�д����ɡ��㽣�󡿣�https://www.cuijiange.com����Ʒ��ּ����߹���Ч�ʣ������ظ��Ͷ�
'     ������Ȩ��������ο��˲������ѵĹ��ܺ��Ĵ���ʵ�֣���������������˲��ֹ��ܺʹ�����ʹ���߿��Բ������ƵĽ��п������޸��Լ���ҵʹ��
' ���ߣ��㽣�����-���꽣��
' ʹ�÷�����
'     ��Excel�����½����հ׹�����������Alt+F11���򿪡�Visual Basic for Application�༭����
'     �򿪡��ļ�����ѡ�񡾵����ļ��������������������λ�ã�ѡ�б��ļ��������ȷ��������
'     �رա�Visual Basic for Application�༭��������������Ϊ���������ļ�����Ϊ��Excel���غ�|*.xlam�������ȷ��

' ģ�����ƣ��㽣���ģ��_20200221
' ģ�鹦�ܣ�
' 1���Զ������иߣ�

' 2���Զ�����ͼƬ��С��

' 3����ָ���ַ��ͼ���ָ��ı�(splitIn)

Sub �Զ������и�()
'     ��ݷ�ʽ����
'     ����������
'         �Զ�����ѡ�еĵ�Ԫ���иߣ�֧�ֺϲ���Ԫ��֧��ͬʱѡ������Ԫ������
'     ʹ��Ҫ�죺
'         ѡ��һ��������Ҫ�����иߵĵ�Ԫ��ִ�к�
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
        g = MsgBox("����ѡ����Ҫ'������и�'����!", vbInformation)
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
    Application.DisplayAlerts = False            'ɾ������������ʾȥ��
    wrkSheet.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
Public Sub �Զ�����ͼƬ��С()
Attribute �Զ�����ͼƬ��С.VB_ProcData.VB_Invoke_Func = "Q\n14"

'     ��ݷ�ʽ��Ctrl+Shift+Q���粻��Ҫ��ݷ�ʽ�����·� Attribute �Զ�����ͼƬ��С.VB_ProcData.VB_Invoke_Func = "Q\n14" ɾ��
'     ����������
'         �Զ��������������ѡ�е�ͼ�δ�С����δѡ���κ�ͼ�Σ���Ĭ�ϵ�������ͼ�Σ���
'         ʹ����Ӧͼ�����Ͻ�����Ӧ�ĵ�Ԫ��Ŀ�ߣ�����ʱ����ԭ��߱Ȳ���
'     ʹ��Ҫ�죺
'         ѡ�е���������Ҫ������С��ͼƬ��ͼ�Σ�ִ�к꣬�粻ѡ����Լ����������ͼ�ν��е���
'Attribute �Զ�����ͼƬ��С.VB_ProcData.VB_Invoke_Func = "Q\n14"

    If TypeName(Selection) = "Range" Then
        ActiveSheet.Shapes.SelectAll
        If TypeName(Selection) = "Range" Then
            MsgBox "���������޿ɵ�������"
            Exit Sub
        End If
    End If
ChangePic Selection
End Sub
Private Sub ChangePic(Optional obj)
'
' ChangePic ��

    Dim s, typ

    If obj Is Nothing Then
        Set s = Selection
        Else
        Set s = obj
        End If
    Dim scaleValue
    scaleValue = 0.95   '1��scaleValue��0
    On Error Resume Next
        If s.TopLeftCell = 0 Then
        End If                  '�ж��Ƿ���ͼ��
        If Err.Number = 0 Then
            typ = "Shape"
        Else
            Err.Clear
            If s.Count <> 0 Then
            End If              '�ж��Ƿ��Ǽ���
            If Err.Number = 0 Then
                typ = "Arrow"
            Else
                MsgBox "��ѡ�񡾵������򡾶����ͼƬ����״"
                Exit Sub
            End If
            
        End If
        Err.Clear
    On Error GoTo 0
    If typ = "Shape" Then
        Dim PicHpW As Double
        Dim RngHpW As Double
        Dim rng As Range
        With s  '����
            
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
'ʹ��ָ���ָ�����ָ����������ַ������зָ���طָ����ַ���
'ʾ����
    'splitIn("1234567890",4," ") ����ֵ��"1234 5678 90"
    '
'str                ��ѡ    ��Ҫ���зָ���ַ���
'SplitCharsNumber   ��ѡ    ��Ҫ�ָ�ļ����    Ĭ��ֵ  4
'SplitChar          ��ѡ    �ָ�ʱ��ʹ�õ��ַ�  Ĭ��ֵ  VbCrLf���س����з���
Function splitIn(ByVal str As String, Optional splitCharsNumber As Integer = 4, Optional splitChar As String = vbCrLf) As String
    Dim RowNumber As Integer, StrLen As Integer
    StrLen = Len(str)
    If splitCharsNumber <= 0 Then
        Err.Raise vbObjectError + 513, , "�ָ�������Ϊ0��ֵ"
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

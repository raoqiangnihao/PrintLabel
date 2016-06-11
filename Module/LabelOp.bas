Attribute VB_Name = "LabelOp"
Option Explicit
Private Const SHT_LABEL As String = "±Í«©"
Private Const SPLIT_CHR As String = "-"
Private Const COL_FINISHED As Long = 4
Sub Label_Init()
    Call InitANewSht(gBk, SHT_LABEL, True)
    Call SaveLstNum("")
End Sub
Sub Label_Print(arr)
    Dim wkSht As Worksheet
    Dim curRow As Long, LstRow As Long, i As Long
    Dim str As String
    Set wkSht = gBk.Worksheets(SHT_LABEL)
    LstRow = wkSht.Cells(wkSht.Rows.count, 1).End(xlUp).Row

    curRow = IIf(LstRow = 1, LstRow, LstRow + 1)
    str = GetNextNum
    wkSht.Cells(curRow, 1) = ExcelApp.WorksheetFunction.Transpose(str)
    Call VPP(curRow)
    wkSht.Cells(curRow, 1).Resize(UBound(arr) - LBound(arr) + 1, 1) = ExcelApp.WorksheetFunction.Transpose(arr)
    Call SaveLstNum(str)
    wkSht.Columns(1).AutoFit
    Set wkSht = Nothing
End Sub
Sub Label_PrintFinish(arr, count As Long)
    Dim wkSht As Worksheet
    Dim LstRow As Long, RowCnt As Long, ColCnt As Long
    Set wkSht = gBk.Worksheets(SHT_LABEL)
    LstRow = wkSht.Cells(wkSht.Rows.count, COL_FINISHED).End(xlUp).Row
    If LstRow <> 1 Then
        Call VPP(LstRow)
    End If
    RowCnt = UBound(arr, 1) - LBound(arr, 1) + 1
    ColCnt = UBound(arr, 2) - LBound(arr, 2) + 1
    wkSht.Cells(LstRow, COL_FINISHED).Resize(RowCnt, ColCnt) = arr
    LstRow = LstRow + RowCnt
    wkSht.Cells(LstRow, COL_FINISHED) = "∞¸ ˝£∫"
    wkSht.Cells(LstRow, COL_FINISHED + 1) = count
    wkSht.Columns.AutoFit
    Set wkSht = Nothing
End Sub

Private Function GetNextNum() As String
    Dim str As String
    Dim arr
    str = GetSetting("PrintLabel", "Label", "Num", "")
    If str = "" Then
        ReDim arr(0 To 1)
        arr(0) = 0
        arr(1) = 1
    Else
        arr = VBA.Split(str, SPLIT_CHR)
        arr(1) = arr(1) + 1
        If arr(1) > 10 Then
            arr(0) = arr(0) + 1
            arr(1) = 1
        End If
    End If
    GetNextNum = VBA.Join(arr, SPLIT_CHR)
End Function
Private Sub SaveLstNum(str As String)
    Call SaveSetting("PrintLabel", "Label", "Num", str)
End Sub


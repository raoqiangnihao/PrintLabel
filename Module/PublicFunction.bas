Attribute VB_Name = "PublicFunction"
Option Explicit
Public Function VPP(ByRef value As Long) As Long
    VPP = value
    value = value + 1
End Function
Public Function BkIsOpen(bkName As String) As Boolean
    Dim wkBk As Workbook
    BkIsOpen = False
    For Each wkBk In ExcelApp.Workbooks
        If wkBk.Name = bkName Then
            BkIsOpen = True
            Exit For
        End If
    Next
    Set wkBk = Nothing
End Function
Public Function HasSht(wkBk As Workbook, ByVal shtName As String) As Boolean
    Dim wkSht As Worksheet
    Dim bRet As Boolean
    bRet = False
    For Each wkSht In wkBk.Worksheets
        If wkSht.Name = shtName Then
            bRet = True
            Exit For
        End If
    Next
    HasSht = bRet
    Set wkSht = Nothing
End Function
Public Function GetSelOpenFiles()
    Dim ofd As FileDialog
    Dim ArrPath, i As Long
    Dim str
    Set ofd = ExcelApp.FileDialog(msoFileDialogFilePicker)
    With ofd
        .AllowMultiSelect = True
        .title = "Select files to Import"
        .Filters.add "Excel文档", "*.xls;*.xlsx", 1
    End With
    If ofd.Show = -1 Then
        ReDim ArrPath(0 To ofd.SelectedItems.count - 1)
        i = 0
        For Each str In ofd.SelectedItems
            ArrPath(VPP(i)) = CStr(str)
        Next
    End If
    Set ofd = Nothing
    GetSelOpenFiles = ArrPath
End Function
Public Function DeleteSht(wkBk As Workbook, ByVal shtName As String) As Boolean
    If HasSht(wkBk, shtName) Then
        ExcelApp.DisplayAlerts = False
        wkBk.Worksheets(shtName).Delete
        ExcelApp.DisplayAlerts = True
        If Err.Number > 0 Then
            DeleteSht = False
            Err.Clear
        Else
            DeleteSht = True
        End If
    End If
End Function
Public Function ArrIsAllValue(arr, value) As Boolean
    Dim i As Long, j As Long
    ArrIsAllValue = True
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            If arr(i, j) <> value Then
                ArrIsAllValue = False
                Exit Function
            End If
        Next
    Next
End Function
Public Sub ActivateSht(wkBk As Excel.Workbook, ByVal shtName)
    If HasSht(wkBk, shtName) Then
        wkBk.Worksheets(shtName).Select
    End If
End Sub
Public Sub InitANewSht(wkBk As Excel.Workbook, shtName As String, bVisible As Boolean)
    Dim wkSht As Worksheet
    Call DeleteSht(wkBk, shtName)
    Set wkSht = wkBk.Worksheets.add(After:=wkBk.Worksheets(wkBk.Worksheets.count))
    wkSht.Name = shtName
    If Not bVisible Then
        wkSht.Visible = Excel.xlSheetHidden
    End If
    Set wkSht = Nothing
End Sub
Public Function PreImport(shtName As String) As Boolean
    If shtName = SHT_SAMPLE Then
        PreImport = False
        Exit Function
    End If
    If shtName = SHT_LABEL Then
        PreImport = False
        Exit Function
    End If
    PreImport = True
End Function
Public Function GerArrLen(arr, d) As Long
    GerArrLen = UBound(arr, d) - LBound(arr, d) + 1
End Function
'******************************************************************************
'函数名称：Sht_GetLstRow
'函数描述：获取工作表的纵向最后非空单元格的行号
'参数说明：
'   wkBk:需要操作的工作表
'   ArgCols：列号
'返回值：
'   最后非空单元格的行号
'******************************************************************************
Public Function Sht_GetLstRow(wkSht As Worksheet, ParamArray ArgCols()) As Long
    Dim i As Long, nRow As Long, nCol As Long, LstRow As Long
    For i = LBound(ArgCols) To UBound(ArgCols)
        nCol = ArgCols(i)
        nRow = wkSht.Cells(wkSht.Rows.count, nCol).End(xlUp).Row
        LstRow = IIf(nRow > LstRow, nRow, LstRow)
    Next
    Sht_GetLstRow = LstRow
End Function

'******************************************************************************
'函数名称：Sht_GetLstRow
'函数描述：获取工作表的横向最后非空单元格的列号
'参数说明：
'   wkBk:需要操作的工作表
'   ArgCols：行号
'返回值：
'   最后非空单元格的列号
'******************************************************************************
Public Function Sht_GetLstCol(wkSht As Worksheet, ParamArray ArgRows()) As Long
    Dim i As Long, nRow As Long, nCol As Long, LstCol As Long
    For i = LBound(ArgRows) To UBound(ArgRows)
        nRow = ArgRows(i)
        nCol = wkSht.Cells(nRow, wkSht.Columns.count).End(xlToLeft).Column
        LstCol = IIf(nCol > LstCol, nCol, LstCol)
    Next
    Sht_GetLstCol = LstCol
End Function

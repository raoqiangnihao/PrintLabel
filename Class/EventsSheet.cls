VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EventsSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const COL_CUR As Long = 2
Private WithEvents EventSheet As Excel.Worksheet
Attribute EventSheet.VB_VarHelpID = -1
Private Sub Class_Terminate()
    If Not EventSheet Is Nothing Then
        Set EventSheet = Nothing
    End If
End Sub
Public Sub SetEventSheet(wkSht As Excel.Worksheet)
    Set EventSheet = wkSht
End Sub
Public Sub SetValidation(arrPaths)
    Dim Rng As Range
    Dim i As Long, str As String
    For i = LBound(arrPaths) To UBound(arrPaths)
        If arrPaths(i) <> "" Then
            str = str & Dir(arrPaths(i)) & ","
        End If
    Next
    str = Left(str, Len(str) - 1)
    Set Rng = EventSheet.Cells(1, COL_CUR)
    ExcelApp.EnableEvents = False
    Rng.Validation.Delete
    Rng.Validation.Add xlValidateList, xlValidAlertStop, xlBetween, str
    Rng = Dir(arrPaths(LBound(arrPaths)))
    Rng.Offset(0, -1) = "当前处理工作簿"
    EventSheet.Columns.AutoFit
    ExcelApp.EnableEvents = True
End Sub
Public Function GetCurHandle() As String
    GetCurHandle = EventSheet.Cells(1, COL_CUR)
End Function
Public Sub ScanInit()
    Dim nRow As Long, nCol As Long
    ExcelApp.EnableEvents = False
    For nRow = EventSheet.UsedRange.Rows.count To 2 Step -1
        EventSheet.Rows(nRow).Delete
    Next
    For nCol = EventSheet.UsedRange.Columns.count To COL_CUR + 1 Step -1
        EventSheet.Columns(nCol).Delete
    Next
    EventSheet.Activate
    EventSheet.Cells(2, 1).Select
    ExcelApp.EnableEvents = True
End Sub


Private Sub EventSheet_Change(ByVal Target As Excel.Range)
    On Error Resume Next
    ExcelApp.EnableEvents = False
    If Target.Column <= COL_CUR And Target.Row = 1 Then
        If Target.Column = COL_CUR Then
            Dim RowCnt As Long, ColCnt As Long
            Dim arr
            Call gShtScan.ScanInit
            ExcelApp.EnableEvents = False
            arr = Sample_GetScannedCode
            If IsArray(arr) Then
                RowCnt = GerArrLen(arr, 1)
                ColCnt = GerArrLen(arr, 2)
                Target.Offset(1, -1).Resize(RowCnt, ColCnt) = arr
            Else
                RowCnt = 1
                Target.Offset(1, -1) = arr
            End If
            Target.Offset(RowCnt + 1, -1).Select
        End If
        Exit Sub
    End If
    
    If True Then
        If Target.Cells.count = 1 Then
            Call Sample_AddScanResult(Target.value)
        End If
    End If
    ExcelApp.EnableEvents = True
End Sub


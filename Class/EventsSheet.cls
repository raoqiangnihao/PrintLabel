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

'函数描述：类的构造函数
Private Sub Class_Terminate()
    If Not EventSheet Is Nothing Then
        Set EventSheet = Nothing
    End If
End Sub

'函数描述：设置需要监视的工作表
'参数说明：
'   wkSht   需要监视的工作表
Public Sub SetEventSheet(wkSht As Excel.Worksheet)
    Set EventSheet = wkSht
End Sub

'函数描述：设置正在扫描工作簿的数据有效性
'参数说明：
'   arrPaths    有效性内容，数组
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
    EventSheet.Columns(1).ColumnWidth = COL_WIDTH_CODE
    'EventSheet.Columns(1).HorizontalAlignment = xlCenter
    EventSheet.Columns(2).AutoFit
    ExcelApp.EnableEvents = True
End Sub

'函数描述：获取当前正在扫描的工作簿的名称
Public Function GetCurHandle() As String
    GetCurHandle = EventSheet.Cells(1, COL_CUR)
End Function

'函数描述：初始化扫描结果工作表
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
    EventSheet.Cells(2, 1).Select '默认开始位置为A2
    ExcelApp.EnableEvents = True
End Sub

'工作表的Change事件
Private Sub EventSheet_Change(ByVal Target As Excel.Range)
    On Error Resume Next
    ExcelApp.EnableEvents = False
    If Target.Column <= COL_CUR And Target.Row = 1 Then
        If Target.Column = COL_CUR Then
            '如果选择的扫描工作簿，则将以前扫描的结果搬移过来
            Dim RowCnt As Long, ColCnt As Long
            Dim arr
            Call gShtScan.ScanInit
            ExcelApp.EnableEvents = False
            arr = Sample_GetScannedCode
            If IsArray(arr) Then
                '扫描多个
                RowCnt = GerArrLen(arr, 1)
                ColCnt = GerArrLen(arr, 2)
                Target.Offset(1, -1).Resize(RowCnt, ColCnt) = arr
            Else
                If arr <> "" Then
                    RowCnt = 1
                    Target.Offset(1, -1) = arr '已经扫描了一个
                Else
                    RowCnt = 0 '如果一个都没有扫描
                End If
            End If
            Target.Offset(RowCnt + 1, -1).Select '选择下一个输入的位置
        End If
        GoTo LineEnd
    End If
    
    If IsScanning Then
        If Target.Cells.count = 1 Then
            If Trim(Target.value) <> "" Then
                If Not Sample_AddScanResult(Target.value) Then
                    Target.ClearContents
                End If
            End If
        End If
    End If
LineEnd:
    ExcelApp.EnableEvents = True
End Sub


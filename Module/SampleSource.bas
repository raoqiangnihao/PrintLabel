Attribute VB_Name = "SampleSource"
Option Explicit
Private Const SHT_SAMPLE As String = "src_sample"
Private Const SYMBOL_END As String = "end"
Private Const COl_DEPART As Long = 4
Private Const COL_SCANNED As Long = 3
Private Const ROW_SCANNED As Long = 6

Sub Sample_Init()
    Call InitANewSht(gBk, SHT_SAMPLE, True)
End Sub

'函数名称：Sample_ImportData
'功能描述：导入样本
'返回值：true导入成功
Public Function Sample_ImportData() As Boolean
    Dim Paths
    Paths = GetOpenFiles
    If Not IsArray(Paths) Then
        Sample_ImportData = False
        Exit Function
    End If
    Sample_ImportData = True
    Call Sample_Init
    Dim wkBk As Workbook
    Dim i As Long, strPath As String, symbol As String, str As String
    Dim ShtSrc As Worksheet, ShtDst As Worksheet
    Dim Rng As Range
    Dim nCol As Long, nRow As Long, CurCol As Long, CurRow As Long, LstRow As Long
    symbol = "正面条码"
    Set ShtDst = gBk.Worksheets(SHT_SAMPLE)
    CurCol = 1
    For i = LBound(Paths) To UBound(Paths)
        strPath = Paths(i)
        Set wkBk = ExcelApp.Workbooks.Open(strPath)
        For Each ShtSrc In wkBk.Worksheets
            Set Rng = ShtSrc.Rows(4).Find(What:=symbol, LookAt:=xlWhole)
            If Not Rng Is Nothing Then
                CurRow = 1
                ShtDst.Cells(VPP(CurRow), CurCol) = wkBk.Name
                ShtDst.Cells(CurRow, CurCol) = "订单编号："
                ShtDst.Cells(VPP(CurRow), CurCol + 1) = ShtSrc.Cells(2, "B")
                ShtDst.Cells(CurRow, CurCol) = "经销店面："
                ShtDst.Cells(VPP(CurRow), CurCol + 1) = ShtSrc.Cells(3, "B")
                ShtDst.Cells(CurRow, CurCol) = "产品类别："
                str = ShtSrc.Cells(3, "K")
                ShtDst.Cells(VPP(CurRow), CurCol + 1) = Trim(IIf(Len(str) > 5, Right(str, Len(str) - 5), str))
                ShtDst.Cells(CurRow, CurCol + 0) = "扫描订单"
                ShtDst.Cells(CurRow, CurCol + 1) = "样板名称"
                ShtDst.Cells(CurRow, CurCol + 2) = "是否扫描"
                ShtDst.Cells(CurRow, CurCol + 3) = "已经扫描"
                ShtDst.Cells(CurRow + 1, CurCol + 3) = SYMBOL_END
                Call VPP(CurRow)
                nCol = Rng.Column
                
                LstRow = ShtSrc.Cells(ShtSrc.Rows.count, nCol).End(xlUp).Row
                For nRow = Rng.Row + 1 To LstRow
                    ShtDst.Cells(CurRow, CurCol) = ShtSrc.Cells(nRow, nCol)
                    ShtDst.Cells(CurRow, CurCol + 1) = ShtSrc.Cells(nRow, 1)
                    Call VPP(CurRow)
                Next
                CurCol = CurCol + COl_DEPART
            End If
            Set Rng = Nothing
        Next
        wkBk.Close False
        Set ShtSrc = Nothing
        Set wkBk = Nothing
    Next
    ShtDst.Columns.AutoFit
    Set ShtDst = Nothing
    Call gShtScan.SetValidation(Paths)
End Function

'功能描述：获取正在处理的工作簿已经扫描结果
'参数说明：无
'返回值：为空则没有结果，不为空但是不是数组则只有一条结果，是数组则有多条结果，二维数组
Public Function Sample_GetScannedCode()
    Dim wkSht As Worksheet
    Dim CurCol As Long, LstRow As Long, CurRow As Long
    Dim arrRet
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    LstRow = wkSht.Cells(wkSht.Rows.count, CurCol + COL_SCANNED).End(xlUp).Row
    CurRow = ROW_SCANNED + 1
    If LstRow > ROW_SCANNED Then
        arrRet = wkSht.Cells(CurRow, CurCol + COL_SCANNED).Resize(LstRow - CurRow + 1, 1)
    End If
    Set wkSht = Nothing
    Sample_GetScannedCode = arrRet
End Function

'功能描述：添加扫描结果
'参数说明：
'       str 需要添加的条码
'返回值：true当前条码添加成功，false添加失败
Public Function Sample_AddScanResult(str As String) As Boolean
    Dim bRet As Boolean
    Dim wkSht As Worksheet
    Dim LstRow As Long
    Dim CurCol As Long, ScanCol As Long
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    If CurCol = 0 Then
        ShowMsg "先选择一个需要处理的工作簿名称"
        bRet = False
        GoTo LineEnd '不存在当前处理，扫描结果里面没有填当前处理的工作簿
    End If
    ScanCol = CurCol + COL_SCANNED
    If VBA.LCase(str) <> SYMBOL_END Then
        '检测样本中是否包含需要扫描的条码
        If Not ScanCodeIsExist(wkSht, str, CurCol) Then
            ShowMsg "当前处理工作簿，不包含扫描条码，请重新扫描其他条码"
            bRet = False
            GoTo LineEnd
        End If
        '检测是否已经扫描过了
        If ScanCodeIsExist(wkSht, str, ScanCol) Then
            MsgBox "已经扫描过当前条码，请扫描其他的条码"
            bRet = False
            GoTo LineEnd
        End If
    End If
    
    LstRow = wkSht.Cells(wkSht.Rows.count, ScanCol).End(xlUp).Row
    Call VPP(LstRow)
    wkSht.Cells(LstRow, ScanCol) = str
    wkSht.Columns(ScanCol).AutoFit
    Dim bPrint As Boolean, bFinished As Boolean
    bPrint = False: bFinished = False
    If VBA.LCase(str) = SYMBOL_END Then
        '如果是出现了end，则打印标签
        bPrint = True
    Else
        '如果不是end，需要检测是否扫描完
        If CheckFinished(wkSht, str, CurCol) Then
            bPrint = True
            bFinished = True
            Call VPP(LstRow) '如果是扫描完成，则需要下移一行
            Dim msg As String
            msg = "已经扫描完一个工作簿" & Chr(10) & _
                "工作簿：" & wkSht.Cells(1, CurCol) & Chr(10)
            Call ShowMsg(msg)
        End If
    End If
    If bPrint Then
        Dim ArrLabel
        Dim LabelCount As Long
        ArrLabel = GetDisCode(wkSht, LstRow, ScanCol)
        If IsArray(ArrLabel) Then
            ArrLabel = GetDisLabel(ArrLabel, wkSht, CurCol)
            LabelCount = UBound(ArrLabel) - LBound(ArrLabel) + 1
            Call Label_Print(ArrLabel)
            If bFinished Then
                ArrLabel = wkSht.Cells(2, CurCol).Resize(3, 2)
                Call Label_PrintFinish(ArrLabel, LabelCount)
            End If
        End If
    End If
    bRet = True
LineEnd:
    Sample_AddScanResult = bRet
    Set wkSht = Nothing
End Function

'功能描述：设置当前条码的状态，并判断整个工作簿是否扫描完成
'参数说明：
'   wkSht   需要处理的工作表
'   str     需要处理的条码
'   nCol    条码所在的列号
'返回值：true扫描完成，false没有完成
Private Function CheckFinished(wkSht As Worksheet, str As String, ByVal nCol As Long) As Boolean
    Dim Rng As Range
    Dim CurCol As Long, LstRow As Long
    CheckFinished = False
    Set Rng = wkSht.Columns(nCol).Find(What:=str, LookAt:=xlWhole)
    If Not Rng Is Nothing Then
        Rng.Offset(0, 2) = True
        LstRow = wkSht.Cells(wkSht.Rows.count, Rng.Column).End(xlUp).Row
        Dim arr
        arr = wkSht.Range(wkSht.Cells(6, nCol + 2), wkSht.Cells(LstRow, nCol + 2))
        If IsArray(arr) Then
            If ArrIsAllValue(arr, True) Then
                CheckFinished = True
            End If
        Else
            CheckFinished = arr = True
        End If
        Set Rng = Nothing
    End If
End Function

'功能描述：获取需要打印的条码
'参数说明：
'   wkSht   需要处理的工作表
'   LstRow  条码所在列的最后一个不为空的单元格行号
'   nCol    条码所在的列号
'返回值：一个二维数组，如果没有需要显示的，则为空，不是数组
Private Function GetDisCode(wkSht As Worksheet, LstRow As Long, nCol As Long)
    Dim nRow As Long
    Dim str As String
    For nRow = LstRow - 1 To 2 Step -1
        str = wkSht.Cells(nRow, nCol)
        If VBA.LCase(str) = SYMBOL_END Then
            Exit For
        End If
    Next
    Call VPP(nRow)
    If LstRow = nRow Then
        GetDisCode = ""
        Exit Function
    End If
    Dim arrRet
    arrRet = wkSht.Cells(nRow, nCol).Resize(LstRow - nRow, 1)
    If Not IsArray(arrRet) Then
        ReDim arrRet(0 To 0, 0 To 0)
        arrRet(0, 0) = wkSht.Cells(nRow, nCol)
    End If
    GetDisCode = arrRet
End Function

'功能描述：获取需要打印的标签
'参数说明：
'   arr     条码数组
'   wkSht   处理的工作表
'   nCol    条码所在的列号
'返回值：一维数组
Private Function GetDisLabel(arr, wkSht As Worksheet, nCol As Long)
    Dim arrRet, str As String
    Dim i As Long
    Dim Rng As Range
    ReDim arrRet(LBound(arr, 1) To UBound(arr, 1))
    For i = LBound(arr, 1) To UBound(arr, 1)
        str = arr(i, LBound(arr, 2))
        Set Rng = wkSht.Columns(nCol).Find(What:=str, LookAt:=xlWhole)
        If Not Rng Is Nothing Then
            arrRet(i) = Rng.Offset(0, 1)
            Set Rng = Nothing
        End If
    Next
    GetDisLabel = arrRet
End Function

'功能描述：判断指定列中条码是否存在
'参数说明：
'   wkSht   处理的工作表
'   strCode 条码
'   nCol    判断的列号
'返回值：true指定列存在指定的条码
Private Function ScanCodeIsExist(wkSht As Worksheet, strCode As String, nCol As Long) As Boolean
    Dim Rng As Range
    Dim bRet As Boolean
    Set Rng = wkSht.Columns(nCol).Find(What:=strCode, LookAt:=xlWhole)
    If Not Rng Is Nothing Then
        bRet = True
    Else
        bRet = False
    End If
    Set Rng = Nothing
    ScanCodeIsExist = bRet
End Function

'功能描述：获取当前正在处理的工作簿扫描结果所在的列号
'参数说明：
'   wkSht   处理的工作表
'返回值：当前正在扫描的工作簿的结果所在的起始列号
Private Function GetCurHandleCol(wkSht As Worksheet)
    Dim CurHandle As String
    Dim Rng As Range
    CurHandle = gShtScan.GetCurHandle
    If CurHandle = "" Then
        GetCurHandleCol = 0
        Exit Function
    End If
    Set Rng = wkSht.Rows(1).Find(What:=CurHandle, LookAt:=xlWhole)
    If Not Rng Is Nothing Then
        GetCurHandleCol = Rng.Column
        Set Rng = Nothing
    Else
        GetCurHandleCol = 0
    End If
End Function


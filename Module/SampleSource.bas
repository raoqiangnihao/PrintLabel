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

'�������ƣ�Sample_ImportData
'������������������
'����ֵ��true����ɹ�
Public Function Sample_ImportData() As Boolean
    Dim Paths
    Paths = GetOpenFiles
    If Not IsArray(Paths) Then
        Sample_ImportData = False 'û��ѡ���κι�����������ʾ����ʧ��
        Exit Function
    End If
    Sample_ImportData = True
    Call Sample_Init
    Dim wkBk As Workbook
    Dim i As Long, strPath As String, symbol As String, str As String
    Dim ShtSrc As Worksheet, ShtDst As Worksheet
    Dim Rng As Range
    Dim nCol As Long, nRow As Long, CurCol As Long, curRow As Long, LstRow As Long
    symbol = "��������"
    Set ShtDst = gBk.Worksheets(SHT_SAMPLE)
    CurCol = 1
    For i = LBound(Paths) To UBound(Paths)
        strPath = Paths(i)
        Set wkBk = ExcelApp.Workbooks.Open(strPath)
        For Each ShtSrc In wkBk.Worksheets
            Set Rng = ShtSrc.Rows(4).Find(What:=symbol, LookAt:=xlWhole)
            If Not Rng Is Nothing Then
                curRow = 1
                ShtDst.Cells(VPP(curRow), CurCol) = wkBk.Name
                ShtDst.Cells(curRow, CurCol) = "������ţ�"
                ShtDst.Cells(VPP(curRow), CurCol + 1) = ShtSrc.Cells(2, "B")
                ShtDst.Cells(curRow, CurCol) = "�������棺"
                ShtDst.Cells(VPP(curRow), CurCol + 1) = ShtSrc.Cells(3, "B")
                ShtDst.Cells(curRow, CurCol) = "��Ʒ���"
                str = ShtSrc.Cells(3, "K")
                ShtDst.Cells(VPP(curRow), CurCol + 1) = Trim(IIf(Len(str) > 5, Right(str, Len(str) - 5), str))
                ShtDst.Cells(curRow, CurCol + 0) = "��������"
                ShtDst.Cells(curRow, CurCol + 1) = "��������"
                ShtDst.Cells(curRow, CurCol + 2) = "�Ƿ�ɨ��"
                ShtDst.Cells(curRow, CurCol + 3) = "�Ѿ�ɨ��"
                ShtDst.Cells(curRow + 1, CurCol + 3) = SYMBOL_END
                Call VPP(curRow)
                nCol = Rng.Column
                
                LstRow = ShtSrc.Cells(ShtSrc.Rows.count, nCol).End(xlUp).Row
                For nRow = Rng.Row + 1 To LstRow
                    str = ShtSrc.Cells(nRow, 1)
                    If InStr(str, "С��") > 0 And ShtSrc.Cells(nRow, 1).MergeCells Then
                        '�������С�ƣ����˳�
                        Exit For
                    End If
                    ShtDst.Cells(curRow, CurCol) = ShtSrc.Cells(nRow, nCol)
                    ShtDst.Cells(curRow, CurCol + 1) = ShtSrc.Cells(nRow, 1)
                    Call VPP(curRow)
                Next
                ShtDst.Columns(CurCol).ColumnWidth = COL_WIDTH_CODE
                ShtDst.Columns(CurCol + 1).AutoFit
                CurCol = CurCol + COl_DEPART
            End If
            Set Rng = Nothing
        Next
        wkBk.Close False
        Set ShtSrc = Nothing
        Set wkBk = Nothing
    Next
    'ShtDst.Columns.AutoFit
    Set ShtDst = Nothing
    Call gShtScan.SetValidation(Paths)
End Function

'������������ȡ���ڴ���Ĺ������Ѿ�ɨ����
'����˵������
'����ֵ��Ϊ����û�н������Ϊ�յ��ǲ���������ֻ��һ����������������ж����������ά����
Public Function Sample_GetScannedCode()
    Dim wkSht As Worksheet
    Dim CurCol As Long, LstRow As Long, curRow As Long
    Dim arrRet
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    LstRow = wkSht.Cells(wkSht.Rows.count, CurCol + COL_SCANNED).End(xlUp).Row
    curRow = ROW_SCANNED + 1
    If LstRow > ROW_SCANNED Then
        arrRet = wkSht.Cells(curRow, CurCol + COL_SCANNED).Resize(LstRow - curRow + 1, 1)
    End If
    Set wkSht = Nothing
    Sample_GetScannedCode = arrRet
End Function

'�������������ɨ����
'����˵����
'       str ��Ҫ��ӵ�����
'����ֵ��true��ǰ������ӳɹ���false���ʧ��
Public Function Sample_AddScanResult(str As String) As Boolean
    Dim bRet As Boolean
    Dim wkSht As Worksheet
    Dim LstRow As Long
    Dim CurCol As Long, ScanCol As Long
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    If CurCol = 0 Then
        ShowMsg "��ѡ��һ����Ҫ����Ĺ���������"
        bRet = False
        GoTo LineEnd '�����ڵ�ǰ����ɨ��������û���ǰ����Ĺ�����
    End If
    ScanCol = CurCol + COL_SCANNED
    If VBA.LCase(str) <> SYMBOL_END Then
        '����������Ƿ������Ҫɨ�������
        If Not ScanCodeIsExist(wkSht, str, CurCol) Then
            ShowMsg "��ǰ����������������ɨ�����룬������ɨ����������"
            bRet = False
            GoTo LineEnd
        End If
        '����Ƿ��Ѿ�ɨ�����
        If ScanCodeIsExist(wkSht, str, ScanCol) Then
            ShowMsg "�Ѿ�ɨ�����ǰ���룬��ɨ������������"
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
        '����ǳ�����end�����ӡ��ǩ
        bPrint = True
    Else
        '�������end����Ҫ����Ƿ�ɨ����
        If CheckFinished(wkSht, str, CurCol) Then
            bPrint = True
            bFinished = True
            Call VPP(LstRow) '�����ɨ����ɣ�����Ҫ����һ��
            Dim msg As String
            msg = "�Ѿ�ɨ����һ��������" & Chr(10) & _
                "��������" & wkSht.Cells(1, CurCol) & Chr(10)
            Call ShowMsg(msg)
        End If
    End If
    If bPrint Then
        Dim ArrLabel
        Dim LabelCount As Long
        '�����Ҫ��ӡ��ǩ���ӡ��ǩ
        ArrLabel = GetDisCode(wkSht, LstRow, ScanCol)
        If IsArray(ArrLabel) Then
            ArrLabel = GetDisLabel(ArrLabel, wkSht, CurCol)
            LabelCount = GetLableCount(wkSht, ScanCol)
            Call Label_Print(ArrLabel)
            If bFinished Then
                '���������ɨ����ɣ����ӡ���յ���ɱ�ǩ
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

'�������������õ�ǰ�����״̬�����ж������������Ƿ�ɨ�����
'����˵����
'   wkSht   ��Ҫ����Ĺ�����
'   str     ��Ҫ���������
'   nCol    �������ڵ��к�
'����ֵ��trueɨ����ɣ�falseû�����
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

'������������ȡ��Ҫ��ӡ������
'����˵����
'   wkSht   ��Ҫ����Ĺ�����
'   LstRow  ���������е����һ����Ϊ�յĵ�Ԫ���к�
'   nCol    �������ڵ��к�
'����ֵ��һ����ά���飬���û����Ҫ��ʾ�ģ���Ϊ�գ���������
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

'������������ȡ��Ҫ��ӡ�ı�ǩ
'����˵����
'   arr     ��������
'   wkSht   ����Ĺ�����
'   nCol    �������ڵ��к�
'����ֵ��һά����
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

'�������������ص�ǰ��ɨ�����
'����˵����
'   wkSht   ���ڹ�����
'   nCol    �������ڵ��к�
'����ֵ����Ч�������
Private Function GetLableCount(wkSht As Worksheet, nCol As Long) As Long
    Dim LstRow As Long, nRow As Long
    Dim count As Long
    count = 0
    LstRow = wkSht.Cells(wkSht.Rows.count, nCol).End(xlUp).Row
    For nRow = ROW_SCANNED To LstRow
        If wkSht.Cells(nRow, nCol) <> SYMBOL_END Then
            VPP count
        End If
    Next
    GetLableCount = count
End Function

'�����������ж�ָ�����������Ƿ����
'����˵����
'   wkSht   ����Ĺ�����
'   strCode ����
'   nCol    �жϵ��к�
'����ֵ��trueָ���д���ָ��������
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

'������������ȡ��ǰ���ڴ���Ĺ�����ɨ�������ڵ��к�
'����˵����
'   wkSht   ����Ĺ�����
'����ֵ����ǰ����ɨ��Ĺ������Ľ�����ڵ���ʼ�к�
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


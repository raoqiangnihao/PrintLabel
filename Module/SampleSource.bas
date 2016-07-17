Attribute VB_Name = "SampleSource"
Option Explicit
Private Const SHT_SAMPLE As String = "src_sample"
Private Const SYMBOL_END As String = "end"
Private Const SYMBOL_OK As String = "ok"
Private Const COl_DEPART As Long = 5
Private Const COL_SCANNED As Long = 3
Private Const ROW_SCANNED As Long = 6

Sub Sample_Init()
    Call InitANewSht(gBk, SHT_SAMPLE, True)
End Sub

'�������ƣ�Sample_ImportData
'������������������
'����˵����Paths��Ҫ������ļ�����
'����ֵ��true����ɹ�
Public Function Sample_ImportData(Paths) As Boolean
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
        If Dir(strPath) <> "" Then
            Set wkBk = ExcelApp.Workbooks.Open(strPath)
            For Each ShtSrc In wkBk.Worksheets
                Set Rng = ShtSrc.Rows(4).Find(What:=symbol, Lookat:=xlWhole)
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
        End If
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
        arrRet = wkSht.Cells(curRow, CurCol + COL_SCANNED).Resize(LstRow - curRow + 1, 2)
    End If
    Set wkSht = Nothing
    Sample_GetScannedCode = arrRet
End Function

'�������������ɨ����
'����˵����
'       str ��Ҫ��ӵ�����
'����ֵ��true��ǰ������ӳɹ���false���ʧ��
Public Function Sample_AddScanResult(ByVal str As String) As Boolean
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
    If VBA.LCase(wkSht.Cells(LstRow, ScanCol)) = VBA.LCase(str) Then
        bRet = True
        GoTo LineEnd
    End If
    Call VPP(LstRow)
    wkSht.Cells(LstRow, ScanCol) = str
    wkSht.Columns(ScanCol).AutoFit
    
    Dim bPrint As Boolean, bFinished As Boolean
    bPrint = False: bFinished = False
    If VBA.LCase(str) = SYMBOL_END Then
        '����ǳ�����end�����ӡ��ǩ
        'bPrint = True
        '����end����ӡ��ǩ���ȴ������ļ�ɨ����ɺ��ٴ�ӡ
    Else
        '�������end����Ҫ����Ƿ�ɨ����
        If CheckFinished(wkSht, str, CurCol) Then
            gShtScan.InitScanInfo
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
        Call PrintAllLabel
    End If
    bRet = True
LineEnd:
    Sample_AddScanResult = bRet
    Set wkSht = Nothing
End Function

'������������ȡ��ǰ��������������Ϣ
'����˵������
'����ֵ��������Ϣ����
Public Function Sample_GetInfo()
    Dim wkSht As Worksheet
    Dim CurCol As Long
    Dim arr
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    arr = wkSht.Cells(2, CurCol).Resize(3, 3)
    Sample_GetInfo = arr
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
    Set Rng = wkSht.Columns(nCol).Find(What:=str, Lookat:=xlWhole)
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
    wkSht.Cells(2, nCol + 2) = IIf(CheckFinished, "ɨ�����", "")
End Function

'������������ȡ��Ҫ��ӡ������
'����˵����
'   wkSht   ��Ҫ����Ĺ�����
'   nCol    �������ڵ��к�
'����ֵ��һ����ά���飬���û����Ҫ��ʾ�ģ���Ϊ�գ���������
Private Function GetDisCode(wkSht As Worksheet, nCol As Long)
    Dim nRow As Long, LstRow As Long
    Dim str As String
    Dim Rng As Range
    Dim arr
    LstRow = Sht_GetLstRow(wkSht, nCol)
    arr = wkSht.Range(wkSht.Cells(ROW_SCANNED + 1, nCol), wkSht.Cells(LstRow, nCol + 1))
    For nRow = LBound(arr, 1) To UBound(arr, 1)
        str = arr(nRow, 1)
        If VBA.LCase(str) <> SYMBOL_END Then
            Set Rng = wkSht.Columns(nCol - COL_SCANNED).Find(What:=str, Lookat:=xlWhole)
            If Rng Is Nothing Then
                arr(nRow, 2) = ""
            Else
                arr(nRow, 2) = Rng.Offset(0, 1)
            End If
        End If
    Next
    GetDisCode = arr
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
        Set Rng = wkSht.Columns(nCol).Find(What:=str, Lookat:=xlWhole)
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
    Set Rng = wkSht.Columns(nCol).Find(What:=strCode, Lookat:=xlWhole)
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
    Set Rng = wkSht.Rows(1).Find(What:=CurHandle, Lookat:=xlWhole)
    If Not Rng Is Nothing Then
        GetCurHandleCol = Rng.Column
        Set Rng = Nothing
    Else
        GetCurHandleCol = 0
    End If
End Function
'��ӡɨ���ļ������б�ǩ
Private Sub PrintAllLabel()
    Dim ArrLabel
    Dim arrCode
    Dim arrName
    Dim str As String, orderSn As String
    Dim index As Long, count As Long, ScanCol As Long, CurCol As Long
    Dim nRow As Long
    Dim wkSht As Worksheet
    
    Set wkSht = gBk.Worksheets(SHT_SAMPLE)
    CurCol = GetCurHandleCol(wkSht)
    ScanCol = CurCol + COL_SCANNED
    orderSn = wkSht.Cells(2, CurCol + 1)
    '��ȡȫ����ɨ�������
    ArrLabel = GetDisCode(wkSht, ScanCol)
    
    'ͳ���м���
    For nRow = LBound(ArrLabel, 1) To UBound(ArrLabel, 1)
        str = ArrLabel(nRow, LBound(ArrLabel, 2))
        If VBA.LCase(str) = SYMBOL_END Then
            VPP count
        End If
        If nRow = UBound(ArrLabel, 1) And VBA.LCase(str) <> SYMBOL_END Then
            VPP count '���һ�����
        End If
    Next
    index = 1
    ReDim arrCode(0) As String
    ReDim arrName(0) As String
    For nRow = LBound(ArrLabel, 1) To UBound(ArrLabel, 1)
        str = ArrLabel(nRow, LBound(ArrLabel, 2))
        If VBA.LCase(str) = SYMBOL_END Then
            '�������end�����ӡ֮ǰ������
            Call Label_Print(arrName, orderSn, index, count)
            ReDim arrCode(0) As String
            ReDim arrName(0) As String
            VPP index
        Else
            '�������
            arrCode(UBound(arrCode)) = str
            wkSht.Cells(ROW_SCANNED + nRow, ScanCol + 1) = "��" & index & "��"
            arrName(UBound(arrName)) = ArrLabel(nRow, LBound(ArrLabel, 2) + 1)
            ReDim Preserve arrCode(LBound(arrCode) To UBound(arrCode) + 1) As String
            ReDim Preserve arrName(LBound(arrName) To UBound(arrName) + 1) As String
        End If
    Next
    str = ArrLabel(UBound(ArrLabel, 1), LBound(ArrLabel, 2))
    If VBA.LCase(str) <> SYMBOL_END Then
        '�������end�����ӡ֮ǰ������
        Call Label_Print(arrName, orderSn, index, count)
        ReDim arrCode(0) As String
        ReDim arrName(0) As String
    End If
End Sub


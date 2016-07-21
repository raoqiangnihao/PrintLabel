Attribute VB_Name = "LabelOp"
Option Explicit
Public Const SHT_LABEL As String = "标签"
Private Const SPLIT_CHR As String = "-"
Private Const COL_FINISHED As Long = 4
Sub Label_Init()
    Call InitANewSht(gBk, SHT_LABEL, True)
    Call SaveLstNum("")
End Sub
Sub Label_Print(arrName, ByVal Sn As String, ByVal productName As String, dstPlace As String, index As Long, totalCount As Long)
    Dim wkSht As Worksheet
    Dim curRow As Long, LstRow As Long, i As Long, nCol As Long
    Dim str As String
    nCol = 1
    Set wkSht = gBk.Worksheets(SHT_LABEL)
    LstRow = Sht_GetLstRow(wkSht, nCol)
    str = wkSht.Cells(LstRow, nCol)
    '如果是第一行，且为空，则表示是第一次打印
    If Not (LstRow = 1 And str = "") Then
        LstRow = LstRow + 3
    End If
    curRow = LstRow
    If curRow <> 1 Then
        wkSht.HPageBreaks.add wkSht.Cells(curRow, nCol)
    Else
        wkSht.VPageBreaks.add wkSht.Cells(curRow, 4)
    End If
    wkSht.Cells(curRow, nCol) = "订单编号"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = Sn
    wkSht.Cells(curRow, nCol + 1).HorizontalAlignment = xlCenter
    wkSht.Rows(curRow).RowHeight = 33
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    VPP curRow
    
    wkSht.Cells(curRow, nCol) = "产品类别"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = productName
    wkSht.Cells(curRow, nCol + 1).HorizontalAlignment = xlCenter
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    wkSht.Rows(curRow).RowHeight = 33
    
    VPP curRow
    wkSht.Cells(curRow, nCol) = "发货地址"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = dstPlace
    wkSht.Cells(curRow, nCol + 1).WrapText = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    wkSht.Rows(curRow).RowHeight = 33
    
    Dim Dic As New Scripting.Dictionary
    For i = LBound(arrName) To UBound(arrName) - 1
        str = arrName(i)
        Dic(str) = Dic(str) + 1
    Next
    Dim keys, items, count As Long
    keys = Dic.keys
    items = Dic.items
    For i = LBound(keys) To UBound(keys)
        VPP curRow
        wkSht.Cells(curRow, nCol).Resize(1, 2).Merge
        wkSht.Cells(curRow, nCol) = keys(i)
        wkSht.Cells(curRow, nCol).HorizontalAlignment = xlCenter
        count = count + items(i)
        wkSht.Cells(curRow, nCol + 2) = items(i)
    Next
    
    VPP curRow
    wkSht.Cells(curRow, nCol + 0) = "第" & index & "包"
    wkSht.Cells(curRow, nCol + 1) = "共" & totalCount & "包"
    wkSht.Cells(curRow, nCol + 2) = "共" & count & "块"
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    
    Dim Rng As Range
    Set Rng = wkSht.Range(wkSht.Cells(LstRow, nCol), wkSht.Cells(curRow, nCol + 2))
    Rng.Borders.LineStyle = xlContinuous
    Rng.Borders(xlEdgeTop).Weight = xlMedium
    Rng.Borders(xlEdgeLeft).Weight = xlMedium
    Rng.Borders(xlEdgeBottom).Weight = xlMedium
    Rng.Borders(xlEdgeRight).Weight = xlMedium
    
    For i = 0 To 2
        wkSht.Columns(nCol + i).ColumnWidth = 15
        wkSht.Columns(nCol + i).HorizontalAlignment = xlCenter
        wkSht.Columns(nCol + i).VerticalAlignment = xlCenter
    Next
    Dim PrintPage As Long
    PrintPage = wkSht.PageSetup.Pages.count
    wkSht.PrintOut From:=PrintPage, To:=PrintPage
End Sub

Sub Label_PrintFinal(ByVal Sn As String, ByVal productName As String, dstPlace As String, totalPackage As Long, totalCount As Long)
    Dim wkSht As Worksheet, str As String
    Dim curRow As Long, LstRow As Long, nCol As Long
    nCol = 1
    Set wkSht = gBk.Worksheets(SHT_LABEL)
    LstRow = Sht_GetLstRow(wkSht, nCol) + 3
    curRow = LstRow
    wkSht.HPageBreaks.add wkSht.Cells(curRow, nCol)
    
    wkSht.Cells(curRow, nCol) = "订单编号"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = Sn
    wkSht.Cells(curRow, nCol + 1).HorizontalAlignment = xlCenter
    wkSht.Rows(curRow).RowHeight = 33
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    
    VPP curRow
    wkSht.Cells(curRow, nCol) = "产品类别"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = productName
    wkSht.Cells(curRow, nCol + 1).HorizontalAlignment = xlCenter
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    wkSht.Rows(curRow).RowHeight = 33
    
    VPP curRow
    wkSht.Cells(curRow, nCol) = "发货地址"
    wkSht.Cells(curRow, nCol + 1).Resize(1, 2).Merge
    wkSht.Cells(curRow, nCol + 1) = dstPlace
    wkSht.Cells(curRow, nCol + 1).WrapText = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Bold = True
    wkSht.Cells(curRow, nCol).Resize(1, 3).Font.Size = 13
    wkSht.Rows(curRow).RowHeight = 33
    
    VPP curRow
    str = "共 " & totalPackage & " 包 共 " & totalCount & " 块"
    wkSht.Cells(curRow, nCol).Resize(1, 3).Merge
    wkSht.Cells(curRow, nCol) = str
    wkSht.Cells(curRow, nCol).Font.Bold = True
    wkSht.Cells(curRow, nCol).Font.Size = 13
    
    Dim Rng As Range
    Set Rng = wkSht.Range(wkSht.Cells(LstRow, nCol), wkSht.Cells(curRow, nCol + 2))
    Rng.Borders.LineStyle = xlContinuous
    Rng.Borders(xlEdgeTop).Weight = xlMedium
    Rng.Borders(xlEdgeLeft).Weight = xlMedium
    Rng.Borders(xlEdgeBottom).Weight = xlMedium
    Rng.Borders(xlEdgeRight).Weight = xlMedium
    
    Dim PrintPage As Long
    PrintPage = wkSht.PageSetup.Pages.count
    wkSht.PrintOut From:=PrintPage, To:=PrintPage
    
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
    wkSht.Cells(LstRow, COL_FINISHED) = "包数："
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


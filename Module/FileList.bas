Attribute VB_Name = "FileList"
Option Explicit
Private FSO As Scripting.FileSystemObject
'******************************************************************************
'�������ƣ�Fso_Init
'������������ʼ������
'����˵������
'����ֵ����
'******************************************************************************
Private Sub Fso_Init()
    If FSO Is Nothing Then
        Set FSO = New Scripting.FileSystemObject
    End If
End Sub

'******************************************************************************
'������������ȡ�ļ��������ƥ����ļ��б����������ļ���
'����˵����
'   strPath:�ļ���·��
'   pattern:�ļ���ƥ��������ʽ
'����ֵ���ļ��б�����ļ��б�ֻ��һ�������ʾû�л�ȡ�κε��ļ�
'******************************************************************************
Public Function Fso_GetDirFiles(strPath As String, pattern As String)
    Call Fso_Init
    Dim arr
    Dim objFile As Scripting.File
    Dim objFolder As Scripting.Folder
    ReDim arr(0) As String
    Set objFolder = FSO.GetFolder(strPath)
    For Each objFile In objFolder.Files
        If VBA.Len(objFile.Name) > Len(pattern) Then
            If VBA.LCase(VBA.Right(objFile.Name, VBA.Len(pattern))) = VBA.LCase(pattern) Then
                arr(UBound(arr)) = objFile.Path
                ReDim Preserve arr(LBound(arr) To UBound(arr) + 1) As String
            End If
        End If
    Next
    Set objFile = Nothing
    Set objFolder = Nothing
    Fso_GetDirFiles = arr
End Function

'******************************************************************************
'������������ȡ�û�ѡ����ļ�
'����˵����
'   startPath:����Ŀ¼
'   Args���ļ�����
'����ֵ��ѡ����ļ���·��
'******************************************************************************
Public Function Fbd_GetSelFiles(startPath As String, ParamArray args())
    Dim i As Long
    Dim arr
    ReDim arr(0) As String
    Dim ofd As FileDialog
    Set ofd = Application.FileDialog(msoFileDialogFilePicker)
    ofd.Filters.Clear
    For i = LBound(args) To UBound(args) Step 2
        ofd.Filters.add args(i), args(i + 1)
    Next
    ofd.AllowMultiSelect = True
    If startPath <> "" Then
        ofd.InitialFileName = startPath
    End If
    If ofd.Show = -1 Then
        For i = 1 To ofd.SelectedItems.count
            arr(UBound(arr)) = ofd.SelectedItems(i)
            ReDim Preserve arr(LBound(arr) To UBound(arr) + 1) As String
        Next
    End If
    Fbd_GetSelFiles = arr
End Function


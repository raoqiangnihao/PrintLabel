Attribute VB_Name = "FileList"
Option Explicit
Private FSO As Scripting.FileSystemObject
'******************************************************************************
'函数名称：Fso_Init
'函数描述：初始化对象
'参数说明：无
'返回值：无
'******************************************************************************
Private Sub Fso_Init()
    If FSO Is Nothing Then
        Set FSO = New Scripting.FileSystemObject
    End If
End Sub

'******************************************************************************
'函数描述：获取文件夹下面的匹配的文件列表，不包含子文件夹
'参数说明：
'   strPath:文件夹路径
'   pattern:文件名匹配正则表达式
'返回值：文件列表，如果文件列表只有一个，则表示没有获取任何的文件
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
'函数描述：获取用户选择的文件
'参数说明：
'   startPath:初试目录
'   Args：文件类型
'返回值：选择的文件夹路径
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


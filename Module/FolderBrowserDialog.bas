Attribute VB_Name = "FolderBrowserDialog"
Option Explicit
Private Const BIF_STATUSTEXT = &H4&
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SELCHANGED = 2
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_SETSelectION = (WM_USER + 102)

'#If Win64 Then
'Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'#Else
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'#End If

Private Type BrowseInfo
    hWndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

'The current directory
Private m_CurrentDirectory As String

'******************************************************************************
'函数描述：获取用户选择的文件夹路径
'参数说明：
'   title:标题
'   startDir：开始路径
'返回值：选择的文件夹路径
'******************************************************************************
Public Function Fbd_GetSelDir(Optional title As String = "选择文件夹", Optional startDir As String = "") As String
    Dim str As String
    str = BrowseForFolder(Application.hWnd, title, startDir)
    Fbd_GetSelDir = str
End Function

Private Function BrowseForFolder(hWnd As Long, title As String, startDir As String) As String
    Dim lpIDList As Long
    Dim szTitle As String
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo
    m_CurrentDirectory = startDir & vbNullChar

    szTitle = title
    With tBrowseInfo
        .hWndOwner = hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN + BIF_STATUSTEXT
        .lpfnCallback = GetAddressofFunction(AddressOf BrowseCallbackProc)  'get address of function.
    End With
    
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    Else
        BrowseForFolder = ""
    End If
End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long
    Dim lpIDList As Long
    Dim ret As Long
    Dim sBuffer As String
    On Error Resume Next
    Select Case uMsg
        Case BFFM_INITIALIZED
            Call SendMessage(hWnd, BFFM_SETSelectION, 1, m_CurrentDirectory)
        Case BFFM_SELCHANGED
            sBuffer = Space(MAX_PATH)
    
            ret = SHGetPathFromIDList(lp, sBuffer)
            If ret = 1 Then
                Call SendMessage(hWnd, BFFM_SETSTATUSTEXT, 0, sBuffer)
            End If
    End Select
    BrowseCallbackProc = 0
End Function
Private Function GetAddressofFunction(add As Long) As Long
  GetAddressofFunction = add
End Function

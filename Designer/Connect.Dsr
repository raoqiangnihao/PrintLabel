VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   10215
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   15300
   _ExtentX        =   26988
   _ExtentY        =   18018
   _Version        =   393216
   Description     =   "Add-In Project Template"
   DisplayName     =   "My Add-In"
   AppName         =   "Microsoft Excel"
   AppVer          =   "Microsoft Excel 14.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Implements IRibbonExtensibility '添加对 IRibbonExtensibility 接口的引用
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set ExcelApp = Application
    SftVer = App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Set ExcelApp = Nothing
    Set gBk = Nothing
    Set gShtScan = Nothing
End Sub

'调用自定义 XML
Public Function IRibbonExtensibility_GetCustomUI(ByVal RibbonID As String) As String
    IRibbonExtensibility_GetCustomUI = GetRibbonXML()
End Function

'添加 XML 自定义代码
Public Function GetRibbonXML() As String
    Dim sRibbonXML As String
    sRibbonXML = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >" & _
                  "<ribbon startFromScratch=""false"">" & _
                   "<tabs>" & _
                    "<tab id=""tabPrintLabel"" label=""打印标签"">" & _
                     "<group id=""grupPrintLabel"" label=""工具 V " & App.Major & "." & App.Minor & "." & App.Revision & """>" & _
                      "<button id=""btnImportSelSample"" label=""导入选择样本数据"" size=""large"" imageMso=""ExportMoreMenu"" onAction=""UIImportSelSample"" />" & _
                      "<button id=""btnImportFolderSample"" label=""导入文件夹样本数据"" size=""large"" imageMso=""SharePointListsWorkOffline"" onAction=""UIImportFolderSample"" />" & _
                      "<button id=""btnClearScan"" label=""清除扫描数据"" size=""large"" imageMso=""InkDeleteAllInk"" onAction=""UIClearScan"" />" & _
                     "</group >" & _
                    "</tab>" & _
                   "</tabs>" & _
                  "</ribbon>" & _
                 "</customUI>"
    GetRibbonXML = sRibbonXML
End Function
   
Public Sub UIImportSelSample(control As IRibbonControl)
'    On Error Resume Next
    Dim wkSht As Excel.Worksheet
    Set wkSht = ExcelApp.ActiveSheet
    If Not PreImport(wkSht.Name) Then
        ShowMsg "当前工作表不能作为扫描结果，请新建一个工作表"
        Set gShtScan = Nothing
        Set wkSht = Nothing
        Exit Sub
    End If
    Dim arrPaths
    arrPaths = GetSelOpenFiles
    Call UIImport(wkSht, arrPaths)
End Sub
Public Sub UIImportFolderSample(control As IRibbonControl)
    Dim wkSht As Excel.Worksheet
    Set wkSht = ExcelApp.ActiveSheet
    If Not PreImport(wkSht.Name) Then
        ShowMsg "当前工作表不能作为扫描结果，请新建一个工作表"
        Set gShtScan = Nothing
        Set wkSht = Nothing
        Exit Sub
    End If
    Dim SelDir As String
    Dim arrPaths
    SelDir = Fbd_GetSelDir
    arrPaths = Fso_GetDirFiles(SelDir, ".xlsx")
    If UBound(arrPaths) > 0 Then
        ReDim Preserve arrPaths(LBound(arrPaths) To UBound(arrPaths) - 1) As String
        Call UIImport(wkSht, arrPaths)
    End If
    
End Sub

Private Sub UIImport(wkSht As Worksheet, arrPaths)
    If Not gBk Is Nothing Then
        Set gBk = Nothing
    End If
    Set gBk = ExcelApp.ActiveWorkbook
    If Not gShtScan Is Nothing Then
        Set gShtScan = Nothing
    End If
    Set gShtScan = New EventsSheet
    Call gShtScan.SetEventSheet(wkSht)
    Dim bSucess As Boolean
    
    ExcelApp.ScreenUpdating = False
    bSucess = Sample_ImportData(arrPaths)
    If bSucess Then
        Call Label_Init
        Call gShtScan.ScanInit
    End If
    ExcelApp.ScreenUpdating = True
    If bSucess Then
        Call ShowMsg("导入成功")
    Else
        Call ShowMsg("导入失败")
    End If
End Sub



Sub UIClearScan(control As IRibbonControl)
    If Not gShtScan Is Nothing Then
        Call gShtScan.ScanInit
    End If
End Sub


VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   7965
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   14049
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
Implements IRibbonExtensibility '��Ӷ� IRibbonExtensibility �ӿڵ�����
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    Set ExcelApp = Application
End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    Set ExcelApp = Nothing
    Set gBk = Nothing
    Set gShtScan = Nothing
End Sub

'�����Զ��� XML
Public Function IRibbonExtensibility_GetCustomUI(ByVal RibbonID As String) As String
    IRibbonExtensibility_GetCustomUI = GetRibbonXML()
End Function

'��� XML �Զ������
Public Function GetRibbonXML() As String
    Dim sRibbonXML As String
    sRibbonXML = "<customUI xmlns=""http://schemas.microsoft.com/office/2006/01/customui"" >" & _
                  "<ribbon startFromScratch=""false"">" & _
                   "<tabs>" & _
                    "<tab id=""tabPrintLabel"" label=""��ӡ��ǩ"">" & _
                     "<group id=""grupPrintLabel"" label=""����"">" & _
                      "<button id=""btnImportSample"" label=""������������"" size=""large"" imageMso=""ExportMoreMenu"" onAction=""UIImportSample"" />" & _
                      "<button id=""btnClearScan"" label=""���ɨ������"" size=""large"" imageMso=""InkDeleteAllInk"" onAction=""UIClearScan"" />" & _
                     "</group >" & _
                    "</tab>" & _
                   "</tabs>" & _
                  "</ribbon>" & _
                 "</customUI>"
    GetRibbonXML = sRibbonXML
End Function
   
Public Sub UIImportSample(control As IRibbonControl)
'    On Error Resume Next
    Dim wkSht As Excel.Worksheet
    Set wkSht = ExcelApp.ActiveSheet
    If Not PreImport(wkSht.Name) Then
        ShowMsg "��ǰ����������Ϊɨ���������½�һ��������"
        Set gShtScan = Nothing
        Set wkSht = Nothing
        Exit Sub
    End If
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
    bSucess = Sample_ImportData
    If bSucess Then
        Call Label_Init
        Call gShtScan.ScanInit
    End If
    ExcelApp.ScreenUpdating = True
    If bSucess Then
        Call ShowMsg("����ɹ�")
    Else
        Call ShowMsg("����ʧ��")
    End If
End Sub
Sub UIClearScan(control As IRibbonControl)
    If Not gShtScan Is Nothing Then
        Call gShtScan.ScanInit
    End If
End Sub


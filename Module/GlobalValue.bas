Attribute VB_Name = "GlobalValue"
Option Explicit
'标签打印
'Public Const SHT_LABEL As String = "标签"
'是否正在扫描
Public Const IsScanning As Boolean = True
'条码所在的列列宽
Public Const COL_WIDTH_CODE As Long = 20

Public ExcelApp As Excel.Application
Public gBk As Excel.Workbook
Public gShtScan As EventsSheet
Public SftVer As String


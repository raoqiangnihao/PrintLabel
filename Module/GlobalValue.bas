Attribute VB_Name = "GlobalValue"
'样本所在的工作表名称
Public Const SHT_SAMPLE As String = "src_sample"
'标签打印
Public Const SHT_LABEL As String = "标签"
'是否正在扫描
Public Const IsScanning As Boolean = True

Public ExcelApp As Excel.Application
Public gBk As Excel.Workbook
Public gShtScan As EventsSheet


Attribute VB_Name = "GlobalValue"
Option Explicit
'��ǩ��ӡ
'Public Const SHT_LABEL As String = "��ǩ"
'�Ƿ�����ɨ��
Public Const IsScanning As Boolean = True
'�������ڵ����п�
Public Const COL_WIDTH_CODE As Long = 20

Public ExcelApp As Excel.Application
Public gBk As Excel.Workbook
Public gShtScan As EventsSheet
Public SftVer As String


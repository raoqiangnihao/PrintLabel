Attribute VB_Name = "UIOnAction"
Option Explicit
Private Const MSG_TITLE = "����"


Sub ShowMsg(ByVal msg As String)
    MsgBox msg, vbOKOnly, MSG_TITLE
End Sub

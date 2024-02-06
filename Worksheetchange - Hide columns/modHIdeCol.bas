Attribute VB_Name = "modHIdeCol"
Option Explicit

Public Sub HideColumns(number As Long)
Attribute HideColumns.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'
    If number > 8 Then
    ActiveSheet.Columns("D:E").Hidden = False
    Else
    ActiveSheet.Columns("D:E").Hidden = True
    End If
End Sub

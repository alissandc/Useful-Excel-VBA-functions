Attribute VB_Name = "modHIdeCol"
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

If Range("U1").Value > 8 Then

Columns("D:E").Hidden = False
Columns("C:E").EntireColumn.AutoFit
Else

Columns("D:E").Hidden = True
Columns("C:C").ColumnWidth = 40

End If


'    'If Target.Cells.Count > 1 Then GoTo done
'
'    If Application.Intersect(Target, ActiveSheet.Range("U1")) Is Nothing Then GoTo done
'
'    'Application.EnableEvents = False
'
'    HideColumns Target.Value
'
'done:
'    'Application.EnableEvents = True
'    Exit Sub
End Sub




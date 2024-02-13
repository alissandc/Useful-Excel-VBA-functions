Attribute VB_Name = "Module2"
Option Explicit

Public Sub findandreplaceduplicates()
Attribute findandreplaceduplicates.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'

Dim companyname As String, propertyname As String, count As Long
Dim dstcount As Long
Dim replacewith As String
Dim num As Long
Dim skip_row_num As Long
Dim cntr As Long
Dim data_wb As Workbook
Set data_wb = Workbooks.Open("C:\Users\acayabyab\OneDrive - RealPage\Documents\OpsMerchant Top 100 quarterly\property\20240201 PROPERTYTABLEDATA.xlsx")

     For cntr = 101 To 2029
        
        propertyname = ThisWorkbook.Sheets("duplicates").Range("A" & cntr).Value
        companyname = ThisWorkbook.Sheets("duplicates").Range("B" & cntr).Value
        dstcount = ThisWorkbook.Sheets("duplicates").Range("D" & cntr).Value
        replacewith = ThisWorkbook.Sheets("duplicates").Range("E" & cntr).Value
        
        If dstcount = 1 Then GoTo Skip
        
        data_wb.Sheets("Data").Activate
        'data_wb.Sheets("Data").ShowAllData
        data_wb.Sheets("Data").Range("$A$1:$N$788799").AutoFilter Field:=9, Criteria1:=companyname
        data_wb.Sheets("Data").Range("$A$1:$N$788799").AutoFilter Field:=12, Criteria1:=propertyname
        data_wb.Sheets("Data").Range("$A$2:$N$788799").Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Replace What:=propertyname, Replacement:=replacewith, _
        LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:= _
        False, ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        ActiveSheet.ShowAllData
        
Skip:
        
    Next
    
    MsgBox "done"

End Sub

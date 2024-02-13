Attribute VB_Name = "property"
Option Explicit

Public Sub mainProperty2()
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim templatepath As String
    Dim datafilepath As String
    Dim resultfolderpath As String
    templatepath = ThisWorkbook.Sheets("Main").Range("L3").Value
    datafilepath = ThisWorkbook.Sheets("Main").Range("L4").Value
    resultfolderpath = ThisWorkbook.Sheets("Main").Range("L5").Value
    Dim wb_data As Workbook
    Set wb_data = Workbooks.Open(datafilepath)
    
  
    Dim cntr As Long
    Dim suppliername As String
    Dim newfilepath As String, newfilename As String
    Dim lastrow_alldetails As Long, lastrow_newclients As Long, lastrow_existingclients, lastrow_qo As Long
    Dim checknumber As Long
    Dim NewClientSum As Long
    Dim ec_row1 As Long, ec_row2 As Long
    Dim strPassword As String
    strPassword = "878MeRr7Ov33"
    Dim count_new_clients As Long, hide_new_clients As Long, hide_new_clients2 As Long
    Dim lr_newclients As Long
    
    For cntr = 49 To 105
        suppliername = ThisWorkbook.Sheets("Main").Range("A" & cntr).Value
        newfilename = ThisWorkbook.Sheets("Main").Range("B" & cntr).Value
        newfilepath = resultfolderpath & newfilename & ".xlsb"
        
        wb_data.Sheets("Data").Activate
        wb_data.Sheets("Data").Range("$A$1:$N$788799").AutoFilter Field:=5, Criteria1:=suppliername
        checknumber = wb_data.Sheets("Data").Range("P1").Value
        If checknumber < 1 Then
            ThisWorkbook.Sheets("Main").Range("C" & cntr).Value = "No data."
            GoTo skip
        End If
                
        oFSO.CopyFile templatepath, newfilepath
        Dim new_wb As Workbook
        Set new_wb = Workbooks.Open(newfilepath)
        
        new_wb.Sheets("All Details").Range("A2:X2000").Clear
                
        wb_data.Sheets("Data").Activate
        wb_data.Sheets("Data").Range("$A$2:$N$788799").Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        new_wb.Sheets("All Details").Activate
        new_wb.Sheets("All Details").Range("A2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        lastrow_alldetails = new_wb.Sheets("All Details").Cells(Sheets("All Details").Rows.Count, 1).End(xlUp).Row
        new_wb.Sheets("All Details").Sort.SortFields.Clear
        new_wb.Sheets("All Details").Range("A1:X" & lastrow_alldetails).Sort Key1:=Range("A1"), Key2:=Range("I1"), Key3:=Range("L1"), Header:=xlYes, _
        Order1:=xlAscending, Order2:=xlDescending
        
        'formulas
        Application.Calculation = xlManual
        new_wb.Sheets("All Details").Range("O2").Formula = "=YEAR(A2)"
        new_wb.Sheets("All Details").Range("P2").Formula = "=CONCATENATE(""Q"",ROUNDUP(MONTH($A2)/3,0))"
        new_wb.Sheets("All Details").Range("Q2").Formula = "=CONCATENATE(P2,"" "",O2)"
        new_wb.Sheets("All Details").Range("R2").Formula = "=IF(COUNTIFS($I$2:I2,$I2,$Q$2:Q2,Q2)>1,"""",1)"
        new_wb.Sheets("All Details").Range("S2").Formula = "=SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,S$1)"
        new_wb.Sheets("All Details").Range("T2").Formula = "=IF(AND(SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,S$1)<=0,SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,T$1)>0),""New Clients"",""Existing Clients"")"
        new_wb.Sheets("All Details").Range("U2").Formula = "=IF(AND(SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,T$1)<=0,SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,U$1)>0),""New Clients"",""Existing Clients"")"
        new_wb.Sheets("All Details").Range("V2").Formula = "=IF(AND(SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,U$1)<=0,SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,V$1)>0),""New Clients"",""Existing Clients"")"
        new_wb.Sheets("All Details").Range("W2").Formula = "=IF(AND(SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,V$1)<=0,SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,W$1)>0),""New Clients"",""Existing Clients"")"
        new_wb.Sheets("All Details").Range("X2").Formula = "=IF(AND(SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,W$1)<=0,SUMIFS($N:$N,$I:$I,$I2,$Q:$Q,X$1)>0),""New Clients"",""Existing Clients"")"
        
        
        
        new_wb.Sheets("All Details").Range("O2:X2").Select
        Selection.AutoFill destination:=new_wb.Sheets("All Details").Range("O2:X" & lastrow_alldetails), Type:=xlFillDefault
        Application.Calculation = xlAutomatic
        new_wb.Sheets("Quarterly overview").Activate
        new_wb.Sheets("Quarterly overview").PivotTables("NewClients").PivotCache.Refresh
        new_wb.Sheets("Quarterly overview").PivotTables("ExistingClients").PivotCache.Refresh
        
        lastrow_newclients = new_wb.Sheets("Quarterly overview").Range("D1048576").End(xlUp).Row
        lastrow_existingclients = new_wb.Sheets("Quarterly overview").Range("M1048576").End(xlUp).Row
        
        NewClientSum = new_wb.Sheets("Quarterly overview").Range("M14").Value
        
        Application.Calculation = xlManual
        
        If NewClientSum > 0 Then
        'have new clients
            new_wb.Sheets("Quarterly overview").PivotTables("NewClients").PivotFields("Q4 2023").ClearAllFilters
            new_wb.Sheets("Quarterly overview").PivotTables("NewClients").PivotFields("Q4 2023").CurrentPage = _
            "New Clients"
            
            'get new client list
            new_wb.Sheets("All Details").Range("A1:X" & lastrow_alldetails).AutoFilter Field:=17, Criteria1:="Q4 2023"
            new_wb.Sheets("All Details").Range("A1:X" & lastrow_alldetails).AutoFilter Field:=24, Criteria1:="New Clients"
            'new_wb.Sheets("All Details").Columns("I2:I1047592").Copy new_wb.Sheets("All Details").Range("AE1")
            new_wb.Sheets("All Details").Activate
            new_wb.Sheets("All Details").Range("I2:I" & lastrow_alldetails).Select
            Selection.SpecialCells(xlCellTypeVisible).Copy new_wb.Sheets("All Details").Range("AE1")
            new_wb.Sheets("All Details").ShowAllData
            'new_wb.Sheets("All Details").Columns("AE:AE").Select
            new_wb.Sheets("All Details").Range("$AE$1:$AE$1047592").RemoveDuplicates Columns:=1, Header:=xlYes
            count_new_clients = WorksheetFunction.CountA(new_wb.Sheets("All Details").Columns("AE:AE"))
        Else
        'no new clients
            ActiveSheet.PivotTables("NewClients").PivotFields("Q4 2023").ClearAllFilters
            ActiveSheet.PivotTables("NewClients").PivotFields("Q4 2023").CurrentPage = _
            "(blank)"
            new_wb.Sheets("Quarterly overview").Range("C39").Value = "None"
            new_wb.Sheets("Quarterly overview").Range("C39").Font.Italic = True
        End If
        
        If lastrow_newclients < 39 Then
        'no new clients
            new_wb.Sheets("Quarterly overview").Range("C39").Value = "None"
            new_wb.Sheets("Quarterly overview").Range("C39").Font.Italic = True
        Else
            new_wb.Sheets("Quarterly overview").Activate
            new_wb.Sheets("Quarterly overview").Range("E39:F39").Select
            Selection.AutoFill destination:=new_wb.Sheets("Quarterly overview").Range("E39:F" & lastrow_newclients), Type:=xlFillDefault
        End If
        
        If lastrow_existingclients < 39 Then
        'no new clients
            new_wb.Sheets("Quarterly overview").Range("L39").Value = "None"
            new_wb.Sheets("Quarterly overview").Range("L39").Font.Italic = True
        Else
            new_wb.Sheets("Quarterly overview").Activate
            new_wb.Sheets("Quarterly overview").Range("N39:S39").Select
            Selection.AutoFill destination:=new_wb.Sheets("Quarterly overview").Range("N39:S" & lastrow_existingclients), Type:=xlFillDefault
        End If
        
        Application.Calculation = xlAutomatic
        
        lastrow_newclients = new_wb.Sheets("Quarterly overview").Range("F1048576").End(xlUp).Row + 4
        lr_newclients = new_wb.Sheets("Quarterly overview").Range("F1048576").End(xlUp).Row
        
        
        'move existing clients table
        new_wb.Sheets("Quarterly overview").Range("L34:S" & lastrow_existingclients).Select
        Selection.Cut
        new_wb.Sheets("Quarterly overview").Range("C" & lastrow_newclients).Select
        ActiveSheet.Paste
        
        'hide rows
        new_wb.Sheets("Quarterly overview").Rows("35:38").Hidden = True
        ec_row1 = lastrow_newclients + 1
        ec_row2 = lastrow_newclients + 4
        new_wb.Sheets("Quarterly overview").Rows(ec_row1 & ":" & ec_row2).Hidden = True
        
        'remove formulas
        new_wb.Sheets("All Details").Cells.Copy
        new_wb.Sheets("All Details").Cells.PasteSpecial Paste:=xlPasteValues
        
        new_wb.Sheets("Quarterly overview").Range("A4:Q18").Copy
        new_wb.Sheets("Quarterly overview").Range("A4:Q18").PasteSpecial Paste:=xlPasteValues
        
        new_wb.Sheets("Quarterly overview").PivotTables("NewClients").PivotFields("Company Name").ShowDetail = False
        new_wb.Sheets("Quarterly overview").PivotTables("ExistingClients").PivotFields("Company Name").ShowDetail = False
        
        'remove columns in All Details
        new_wb.Sheets("All Details").Columns("S:V").Delete
        new_wb.Sheets("All Details").Columns("O:P").Delete
        new_wb.Sheets("All Details").Columns("K:K").Delete
        new_wb.Sheets("All Details").Columns("D:H").Delete
        new_wb.Sheets("All Details").Columns("B:B").Delete
        
        new_wb.Sheets("Quarterly overview").Columns("L:M").ColumnWidth = 16
        
        If lr_newclients > 50 Then
            new_wb.Sheets("Quarterly overview").Activate
            hide_new_clients = 34 + 4 + count_new_clients + 2
            hide_new_clients2 = ec_row1 - 3
            new_wb.Sheets("Quarterly overview").Range("H32").Value = hide_new_clients
            new_wb.Sheets("Quarterly overview").Range("I32").Value = hide_new_clients2
            new_wb.Sheets("Quarterly overview").Range("G32").Formula = "=COUNTA(D34:D" & lr_newclients & ")"
            new_wb.Sheets("Quarterly overview").Rows(hide_new_clients & ":" & hide_new_clients2).Hidden = True

        End If
        
        new_wb.Sheets("All Details").Cells.Font.Color = vbWhite
        new_wb.Sheets("All Details").Cells.Interior.ColorIndex = 2
        new_wb.Sheets("All Details").Protect Password:=strPassword, DrawingObjects:=True, Contents:=True, Scenarios:=True
        new_wb.Sheets("All Details").EnableSelection = xlNoSelection
        new_wb.Sheets("All Details").Visible = xlSheetVeryHidden
        
        new_wb.Sheets("Quarterly overview").Activate
        lastrow_qo = new_wb.Sheets("Quarterly overview").Range("J1048576").End(xlUp).Row
        new_wb.Sheets("Quarterly overview").Range("E34:J" & lastrow_qo).Locked = True
        new_wb.Sheets("Quarterly overview").Range("E34:J" & lastrow_qo).FormulaHidden = True
        new_wb.Sheets("Quarterly overview").Protect Password:=strPassword, DrawingObjects:=False, Contents:=True, Scenarios:= _
            False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowSorting:=True, AllowFiltering:=True, _
            AllowUsingPivotTables:=True
        new_wb.Sheets("Quarterly overview").EnableSelection = xlNoRestrictions
        new_wb.Sheets("Quarterly overview").Columns("S:XFD").Hidden = True
        new_wb.Sheets("Quarterly overview").Columns("C:C").ColumnWidth = 50
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.Zoom = 100
        new_wb.Sheets("Quarterly overview").Range("A1").Select
        
        wb_data.Sheets("Data").AutoFilter.ShowAllData
        new_wb.Save
        new_wb.Close SaveChanges:=False
        ThisWorkbook.Sheets("Main").Range("C" & cntr).Value = "Done"
        
        
    
    
skip:
    
        'new_wb.Close SaveChanges:=False
    
    Next
    
    'oFSO.CopyFolder "C:\Users\acayabyab\OneDrive - RealPage\Documents\OpsMerchant Top 100 quarterly\property\2023 Q4 Result\", _
    "C:\Users\acayabyab\RealPage\Document Center - PMC Reporting\0006 - eSS Supplier Commissions\Vendor Reports\Merchant Files\"
    'wb_data.Sheets("Data").Activate
    wb_data.Close
    ThisWorkbook.Save
    'MsgBox "done"
End Sub


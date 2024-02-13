Attribute VB_Name = "copyfolder"
Option Explicit

Sub copyfolder()
Attribute copyfolder.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copies all files from a folder to another folder
'

'
    Dim oFSO As Object
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Dim pathsource As String
    Dim destination As String
    Dim oRoot       As Object
    Dim oFile       As Object
    pathsource = "C:\Users\acayabyab\OneDrive - RealPage\Documents\OpsMerchant Top 100 quarterly\property\2023 Q4 Result"
    destination = "C:\Users\acayabyab\RealPage\Document Center - PMC Reporting\0006 - eSS Supplier Commissions\Vendor Reports\Merchant Files\2023 Q4"
    oFSO.copyfolder pathsource, destination

    
    
End Sub

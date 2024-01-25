Attribute VB_Name = "Hider"
Option Explicit

Sub Hide_sheet_in_all_folders()

    Dim folderName As String
    Dim FSOLibrary As Object
    Dim FSOFolder As Object
    Dim FSOFile As Object
    Dim allitems As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select where Results Folder"
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "No file selected. Exitting sub."
            Exit Sub
        Else
            allitems = .SelectedItems(1) & "\"
        End If
    End With

    'Set the file name to a variable
    folderName = allitems

    'Set all the references to the FSO Library
    Set FSOLibrary = CreateObject("Scripting.FileSystemObject")
    Set FSOFolder = FSOLibrary.GetFolder(folderName)

    'Use For Each loop to loop through each file in the folder
    For Each FSOFile In FSOFolder.Files

        'Insert actions to be perfomed on each file
        Dim wbf As Workbook, wbname As String
        wbname = folderName & FSOFile.Name
        Set wbf = Workbooks.Open(wbname)
        wbf.Activate
        wbf.Sheets("All Details").Select
        Application.DisplayAlerts = False
        wbf.Sheets("All Details").Delete
        Application.DisplayAlerts = True
        wbf.Save
        wbf.Close
    Next

    'Release the memory
    Set FSOLibrary = Nothing
    Set FSOFolder = Nothing

End Sub


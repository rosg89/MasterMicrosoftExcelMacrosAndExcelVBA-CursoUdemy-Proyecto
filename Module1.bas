Attribute VB_Name = "Module1"
Public Sub importTextFile()

    Dim TextFile As Workbook
    Dim OpenFiles() As Variant
    Dim i As Integer
    
    OpenFiles = Application.GetOpenFilename(Title:="Select file(s) to import", MultiSelect:=True)
    
    'para que no se abran las ventanas
    Application.ScreenUpdating = False
    
    'for
    For i = 1 To Application.CountA(OpenFiles)
    
        Set TextFile = Workbooks.Open(OpenFiles(i))
        
        TextFile.Sheets(1).Range("A1").CurrentRegion.Copy
        Workbooks(1).Activate
        Workbooks(1).Worksheets.Add
        ActiveSheet.Paste
        'renames the sheet
        ActiveSheet.Name = TextFile.Name
        'clear the clipboard
        Application.CutCopyMode = False
        
        TextFile.Close
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

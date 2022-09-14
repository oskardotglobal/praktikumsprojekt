Attribute VB_Name = "Modul1"
'
' Start Blatt Passwort: 1234
'

' Constants for the EK-VK calculations
Public Const RABATT As Double = 0.885
Public Const HANDLINGKOSTEN As Double = 1.03
Public Const MARGE_CD_LP As Double = 1.6
Public Const MARGE_ANDERE As Double = 1.7
Public Const MINDESTROHERTRAG As Double = 4.5
Public Const MINDESTPREIS As Double = 6.99
Public Const STEUERSATZ As Double = 1.19

' Macro executed by the start button

' Copies the data to the corresponding sheets,
' clears tmp and output, runs preprocessing,
' opens b2b links, runs postprocessing and deselects everything
Sub main()
    ' Show tmp sheet
    Worksheets("tmp").Visible = True
    
    ' Copy data
    Worksheets("Output").Select
    SelectAll
    Selection.delete
    
    Worksheets("tmp").Select
    SelectAll
    Selection.delete
    
    Worksheets("Input").Select
    SelectAll
    Selection.Copy
    Cells(2, 1).Select
    
    Worksheets("tmp").Activate
    Cells(2, 1).Select
    Cells(2, 1).PasteSpecial Paste:=xlPasteValues
    Cells(2, 1).Select
    
    ' Run preprocessing on tmp
    Modul2.PreProcessing
    
    ' Open b2b links in browser if the b2b checkbox is checked
    Dim openB2B
    openB2B = Worksheets("Start").Shapes("openBrowser").ControlFormat.Value
    If openB2B = 1 Then
        OpenLinks
    End If
    
    ' Run postprocessing on Output
    Modul3.PostProcessing
    
    ' Deselect
    Cells(2, 1).Select
    
    ' Hide tmp sheet
    Worksheets("tmp").Visible = False
End Sub

' Macro executed by the "Clear Input" button

' Selects everything in Input except row 1 (captions) and deletes it
Sub ClearInput()
    Worksheets("Input").Activate
    SelectAll
    Selection.delete
    Cells(2, 1).Select
End Sub

' Macro executed by the "Create Import file" button

' Copies the Output sheet into a new file, names the file
' GTG_(the current date in format ddmmyy), opens a folder picker
' for the save location and writes the file
Sub Export()
    Dim name As String
    Dim path As String

    Worksheets("Output").Activate
    Worksheets("Output").Copy
    
    name = "GTG_" & Format(Date, "ddmmyy")
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Wo soll die Importdatei gespeichert werden?"
        .AllowMultiSelect = False
        .InitialFileName = ""
        If .Show <> -1 Then GoTo SaveFile
        path = .SelectedItems(1) & "\"
    End With
    
SaveFile:
        With ActiveWorkbook
            .SaveAs Filename:=path & name, FileFormat:=xlWorkbookNormal, CreateBackup:=False
            .Close False
        End With
End Sub

' Opens all links in column 18 using the file handler by Dev Ashish found in
' Module 4 and deletes the column
Function OpenLinks()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 18).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        Modul4.fHandleFile CStr(cellValue), Modul4.WIN_NORMAL
    Next counter
End Function

' Selects columns A to T and rows 2 to the used range
Function SelectAll()
    If ActiveSheet.UsedRange.Rows.Count < 3 Then
        Range("A2:T" + CStr(3)).Select
    Else
        Range("A2:T" + CStr(ActiveSheet.UsedRange.Rows.Count)).Select
    End If
End Function


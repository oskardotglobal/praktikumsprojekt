Attribute VB_Name = "Modul3"
Sub PostProcessing()
    ' Move columns
    CopyColumn "A", 5
    CopyColumn "B", 10
    CopyColumn "C", 1
    CopyColumn "D", 2
    CopyColumn "E", 13
    CopyColumn "F", 3
    CopyColumn "G", 15
    CopyColumn "H", 4
    CopyColumn "H", 16
    CopyColumn "L", 14
    CopyColumn "S", 11
    CopyColumn "K", 17
    CopyColumn "T", 6
    
    ' Prefill columns
    Prefill "x", 12
    Prefill "Original", 8
    
    ' Calculate VK & EK
    ProcessPrice
    
    ' Calculate year
    ProcessYear
    
    ' Fix formatting of Output
    Worksheets("Output").Activate
    
    ' Color missing
    CheckOutput
    
    Modul1.SelectAll
    
    Selection.NumberFormat = "@"
    With Selection
        .WrapText = True
    End With
    
    ActiveSheet.Columns(13).Select
    Selection.NumberFormat = "dd/mm/yyyy"
    
    ActiveSheet.Columns(15).Select
    Selection.NumberFormat = "0"
    
    ActiveSheet.Columns(9).Select
    Selection.NumberFormat = "0"
End Sub

' Select all rows of the column with letter "source" starting from 2 on sheet "tmp",
' copy them, and paste them on sheet "Output" in the cell 2 of the column number "dest"
Function CopyColumn(source As String, dest As Double)
    Worksheets("tmp").Range(source + "2:" + source + CStr(Worksheets("tmp").UsedRange.Rows.Count)).Select
    Selection.Copy
    Worksheets("Output").Activate
    Cells(2, dest).Select
    ActiveSheet.Paste
    Cells(2, 1).Select
    Worksheets("tmp").Select
    Cells(2, 1).Select
End Function

' Fills all rows of column "columnNumber" with "fill" starting from row 2
Function Prefill(fill As String, columnNumber As Double)
    For counter = 2 To Worksheets("Output").UsedRange.Rows.Count
        If IsEmpty(Worksheets("Output").Cells(counter, 1)) Then
            Exit For
        End If
    
        Worksheets("Output").Cells(counter, columnNumber).Value = fill
    Next counter
End Function

' Takes the hap from from column 17, calculates the EK and VK and fills them
' into column 17 and 18 starting from row 2
Function ProcessPrice()
    For counter = 2 To Worksheets("Output").UsedRange.Rows.Count
        If IsEmpty(Worksheets("Output").Cells(counter, 1)) Then
            Exit For
        End If
    
        Dim cellValue As Double
        Dim ek As Double
        Dim isCDORLP As Boolean
        Dim corresFormat As String
        
        corresFormat = Worksheets("Output").Cells(counter, 5).Value
        
        If InStr(corresFormat, "LP") > 0 Or InStr(corresFormat, "CD") > 0 Then
            isCDORLP = True
        End If
        
        cellValue = Worksheets("Output").Cells(counter, 17).Value
        
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        ek = CalculateEK(cellValue)
        
        Worksheets("Output").Cells(counter, 17).Value = ek
        Worksheets("Output").Cells(counter, 18).Value = CalculateVK(ek, isCDORLP)
    Next counter
End Function

' Get the year from the release date in column 13 and fill it into column 9
Function ProcessYear()
    For counter = 2 To Worksheets("Output").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("Output").Cells(counter, 13).Value
        
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        Worksheets("Output").Cells(counter, 9).Value = Year(cellValue)
        
    Next counter
End Function

' Creates conditional formatting in column 6 and 11 to indicate when a field is empty
Function CheckOutput()
    Cells.FormatConditions.delete
    
    For counter = 2 To Worksheets("Output").UsedRange.Rows.Count
        If IsEmpty(Worksheets("Output").Cells(counter, 1)) Then
            Exit For
        End If
        Dim cellValue
        cellValue = Worksheets("Output").Cells(counter, 6).Value
        
        Worksheets("Output").Cells(counter, 6).Select
            
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=F" & counter & "=" & Chr(34) & Chr(34)
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
        
        cellValue = Worksheets("Output").Cells(counter, 11).Value
        
        Worksheets("Output").Cells(counter, 11).Select
            
        Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=K" & counter & "=" & Chr(34) & Chr(34)
        Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
        With Selection.FormatConditions(1).Interior
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
        End With
    Next counter
End Function

' Helper function to calculate EK
Function CalculateEK(hap As Double) As Double
    CalculateEK = Round(Round(hap * RABATT, 2) * HANDLINGKOSTEN, 2)
End Function

' Helper function to calculate VK
Function CalculateVK(ek As Double, isCDORLP As Boolean)
    Dim vk As Double
    Dim marge As Double
    
    If isCDORLP = True Then
        marge = MARGE_CD_LP
    Else
        marge = MARGE_ANDERE
    End If
    
    vk = Round(Round((ek * marge), 0) - 0.01, 2)
    
    If ((vk / STEUERSATZ) - ek) < 4.5 Then
        vk = Round(Round(((ek + MINDESTROHERTRAG) * STEUERSATZ), 0) - 0.01, 2)
    End If
    
    If (vk < 6.99) Then
        vk = 6.99
    End If
    
    CalculateVK = vk
End Function


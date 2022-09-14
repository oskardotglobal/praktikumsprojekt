Attribute VB_Name = "Modul2"
Sub PreProcessing()
    ' Process data
    DeleteUnwanted 1, 5
    DeleteUnwanted 2, 3
    DeleteUnwanted 6, 1
    
    TrimGenres
    ConvertColumn 2, 10
        
    TrimLabelNames
    TrimArtistNames
    TrimFormats
    TrimTitles
    
    ConvertVGColumn 12, 13
    
    ' Create Format ids in Column 19 by copying the formats and replacing them
    ' using the function ConvertColumn
    Worksheets("tmp").Range("A2:A" + CStr(Worksheets("tmp").UsedRange.Rows.Count)).Select
    Selection.Copy
    Worksheets("tmp").Cells(2, 19).Select
    ActiveSheet.Paste
    ConvertColumn 19, 7
    
    ' Create Country codes in Column 20 by copying the label names and replacing them
    ' using the function ConvertColumn
    Worksheets("tmp").Range("F2:F" + CStr(Worksheets("tmp").UsedRange.Rows.Count)).Select
    Selection.Copy
    Worksheets("tmp").Cells(2, 20).Select
    ActiveSheet.Paste
    ConvertColumn 20, 16
End Sub

' Deletes all rows of column "checkColumn" if it's in the row "sourceColumn"
' on the data sheet, passes the check over to CheckDelete
Function DeleteUnwanted(checkColumn As Double, sourceColumn As Double)
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim condition
        condition = GetDelete(Worksheets("tmp").Cells(counter, checkColumn).Value, sourceColumn)
        If condition = True Then
            Worksheets("tmp").Rows(counter).EntireRow.delete
            counter = counter - 1
        End If
    Next counter
End Function

' Helper function called by DeleteUnwanted, returns true or false depending on
' if the column "columnNumber" on the data sheet contains the string "check"
Function GetDelete(check As String, columnNumber As Double) As Boolean
    For counter = 2 To Worksheets("Stammdaten").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("Stammdaten").Cells(counter, columnNumber).Value
        If IsEmpty(cellValue) Then
            Exit For
        ElseIf InStr(LCase(check), LCase(cellValue)) > 0 Then
        
            Dim genre
            genre = Worksheets("tmp").Cells(counter, 2).Value
            
            If InStr(1, LCase(check), "cd", 0) > 0 And InStr(1, LCase(genre), "hip", 0) > 0 And InStr(1, LCase(genre), "hop", 0) > 0 Then
                Exit For
            End If
            
            GetDelete = True
            Exit For
        End If
    Next counter
End Function

' Replaces all "-" with spaces in the genre column (2)
Function TrimGenres()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 2).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        cellValue = Trim(cellValue)
        cellValue = Replace(cellValue, "-", " ")
        
        Worksheets("tmp").Cells(counter, 2).Value = cellValue
    Next counter
End Function

' Converts the column with number "source",
' passes the string and the column number "data", which is the
' column number of the corresponding columns on the data sheet
' over to the helper function Convert, takes the output and
' fills in the row with that output
Function ConvertColumn(source As Double, data As Double)
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, source).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        Worksheets("tmp").Cells(counter, source).Value = Convert(CStr(cellValue), data)
    Next counter
End Function

' Helper function to convert a column
' Checks if "old" is found in the column with the number sourceColumn on the data sheet,
' if it's found it gets the value from the column to the right side of sourceColumn (sourceColumn + 1),
' strips it of trailing spaces and returns the new value to ConvertColumn
Function Convert(old As String, sourceColumn As Double) As String
    old = Trim(old)

    For counter = 2 To Worksheets("Stammdaten").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Trim(Worksheets("Stammdaten").Cells(counter, sourceColumn).Value)
        If IsEmpty(cellValue) Then
            Exit For
        ElseIf InStr(1, LCase(old), LCase(cellValue), 0) > 0 Then
            Convert = CStr(Worksheets("Stammdaten").Cells(counter, sourceColumn + 1).Value)
            Exit For
        End If
    Next counter
End Function

' Enforces some naming schemes in the label names (column 6),
' like removing "records" from label names if the label name
' isn't "K Records"
Function TrimLabelNames()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 6).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        cellValue = Trim(cellValue)
        
        If InStr(cellValue, "K Records") = 0 Then
            cellValue = Replace(cellValue, "RECORDS", "")
        End If
        
        If InStr(cellValue, "A Recordings") = 0 And InStr(cellValue, "XL Recordings") = 0 Then
            cellValue = Replace(cellValue, "RECORDINGS", "")
        End If
        
        Worksheets("tmp").Cells(counter, 6).Value = cellValue
    Next counter
End Function

' Enforces some naming schemes in the artist names (column 3),
' like removing "OST" from the artist's name and placing it in front of the title
Function TrimArtistNames()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 3).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        If InStr(cellValue, "OST") Then
            titleValue = Worksheets("tmp").Cells(counter, 4).Value
            titleValue = "OST " + titleValue
            Worksheets("tmp").Cells(counter, 4).Value = titleValue
        End If
        
        cellValue = Trim(cellValue)
        cellValue = Replace(cellValue, "OST", "")
        cellValue = Replace(cellValue, "And Others...", "")
        cellValue = Replace(cellValue, "Va.", "Various ")
        cellValue = Replace(cellValue, "/", "")
        
        Worksheets("tmp").Cells(counter, 3).Value = cellValue
    Next counter
End Function

' Enforces some naming schemes in the formats (column 1),
' like removing underscores or replacing information regarding a download
Function TrimFormats()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 1).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        cellValue = Trim(cellValue)
        cellValue = Replace(cellValue, "_", "")
        cellValue = Replace(cellValue, "+MP3", "")
        cellValue = Replace(cellValue, "+ MP3", "")
        cellValue = Replace(cellValue, "+DL", "")
        cellValue = Replace(cellValue, "+ DL", "")
        cellValue = Replace(cellValue, "  ", " ")
        cellValue = Trim(cellValue)
        
        Worksheets("tmp").Cells(counter, 1).Value = cellValue
    Next counter
End Function

' Enforces some naming schemes in the titles (column 4),
' like removing "180g", "Gatefold", "Ltd.", round brackets
' or replacing "Vol." with "Volume"
Function TrimTitles()
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, 4).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        
        cellValue = Trim(cellValue)
        cellValue = Replace(cellValue, "180g", "")
        cellValue = Replace(cellValue, "Gatefold", "")
        cellValue = Replace(cellValue, "Ltd.", "")
        cellValue = Replace(cellValue, "(", "")
        cellValue = Replace(cellValue, ")", "")
        cellValue = Replace(cellValue, "4LP", "")
        cellValue = Replace(cellValue, "+ DL", "")
        cellValue = Replace(cellValue, "+DL", "")
        cellValue = Replace(cellValue, "2LP", "Vinyl Edition")
        cellValue = Replace(cellValue, "Vol.", "Volume ")
        cellValue = Replace(cellValue, "LP", "")
        cellValue = Replace(cellValue, "  ", " ")
        cellValue = Trim(cellValue)
        
        Worksheets("tmp").Cells(counter, 4).Value = cellValue
    Next counter
End Function

' Same as ConvertColumn, but calls ConvertVG instead of Convert
Function ConvertVGColumn(source As Double, data As Double)
    For counter = 2 To Worksheets("tmp").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Worksheets("tmp").Cells(counter, source).Value
        If IsEmpty(cellValue) Then
            Exit For
        End If
        Worksheets("tmp").Cells(counter, source).Value = ConvertVG(CStr(cellValue), data)
    Next counter
End Function

' Same as Convert, but returns "3" if nothing is found
Function ConvertVG(old As String, sourceColumn As Double) As String
    old = Trim(old)

    For counter = 2 To Worksheets("Stammdaten").UsedRange.Rows.Count
        Dim cellValue
        cellValue = Trim(Worksheets("Stammdaten").Cells(counter, sourceColumn).Value)
        If IsEmpty(cellValue) Then
            Exit For
        ElseIf InStr(1, LCase(old), LCase(cellValue), 0) > 0 Then
            ConvertVG = CStr(Worksheets("Stammdaten").Cells(counter, sourceColumn + 1).Value)
            Exit For
        Else
            ConvertVG = "3"
            Exit For
        End If
        
    Next counter
End Function




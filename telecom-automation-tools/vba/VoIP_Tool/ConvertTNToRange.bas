Attribute VB_Name = "Module1"
Sub ConvertTNToRange() 'List to Range
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startTN As Double
    Dim endTN As Double
    Dim currentTN As Double
    Dim nextTN As Double
    Dim outputRow As Long
    Dim colorROw As Boolean
    Dim tnArray() As Double
    Dim headrow As Long
    Dim output As Variant

    ClearResult
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    If lastRow < 2 Then
        MsgBox "Done"
        Exit Sub
    End If
    
    ws.Range("A1:A" & lastRow).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
    
    ' Set the starting range in column D
    outputRow = 2
    startTN = 0
    colorROw = True
    headrow = 1
    
    ReDim tnArray(1 To lastRow - 1) ' Assuming header in row 1
    
    For i = 2 To lastRow
        tnArray(i - 1) = CleanTN(ws.Cells(i, 1).value)
    Next i

    ReDim output(1 To UBound(tnArray), 1 To 6)
    
    ' Loop through each cell in column A
    For i = 1 To UBound(tnArray)
        ' Get the current TN
        currentTN = tnArray(i)
        'set start TN
        If startTN = 0 Then startTN = currentTN

        ' Check if next row is not a number, exit sub
        If i < UBound(tnArray) Then
            nextTN = tnArray(i + 1)
            If Not IsNumeric(nextTN) Then
                MsgBox "A" & i + 1 & " is not a TN"
                Exit Sub
            End If
        Else
            nextTN = 0 ' When reaching the last row
        End If
        
        
        If Not (nextTN = currentTN + 1) Then
        
            If Not (Left(endTN, 3) = Left(currentTN, 3)) Or Not (Mid(endTN, 4, 3) = Mid(currentTN, 4, 3)) Then
                colorROw = Not colorROw
            End If
            
            If colorROw = True Then
                ws.Range("G" & outputRow & ":J" & outputRow).Interior.Color = RGB(228, 228, 228)
            End If
             
            endTN = currentTN ' Set the ending TN
            
            output(outputRow - headrow, 2) = "-->"
            output(outputRow - headrow, 3) = Left(startTN, 3)
            output(outputRow - headrow, 4) = Mid(startTN, 4, 3)
            output(outputRow - headrow, 5) = Right(startTN, 4)
            output(outputRow - headrow, 6) = Right(endTN, 4)
            
             ' Write the range in column D
            If startTN = endTN Then
                output(outputRow - headrow, 1) = startTN
                ws.Cells(outputRow, "J").Font.Bold = True
            Else
                output(outputRow - headrow, 1) = startTN & " to " & endTN
            End If
            'reset startTN
            startTN = 0
            'outputRow to next row
            
            outputRow = outputRow + 1
        End If
    Next i
    
    ws.Range(ws.Cells(headrow + 1, "E"), ws.Cells(outputRow - 1, "J")).value = output
    
    MsgBox "Done"
End Sub


Sub ClearAll()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    If lastRow > 1 Then
        Set rng = ws.Range("A2:J" & lastRow)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"
        ws.Range("E2:E" & lastRow).HorizontalAlignment = xlLeft
        
    End If
    
End Sub

Sub ClearResult()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, "E").End(xlUp).Row
    
    If lastRow > 1 Then
        Set rng = ws.Range("E2:J" & lastRow)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"
        ws.Range("E2:E" & lastRow).HorizontalAlignment = xlLeft
        
    End If
    
End Sub

Function CleanTN(tn As String) As String
    ' Remove -, space, (, )
    tn = Replace(tn, "-", "")
    tn = Replace(tn, " ", "")
    tn = Replace(tn, "(", "")
    tn = Replace(tn, ")", "")
    tn = Replace(tn, ".", "")
    tn = Replace(tn, ",", "")
    
    CleanTN = CDbl(tn)

End Function



Attribute VB_Name = "Module2"
Sub ListTNRange() 'Range to List
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Double
    Dim startTN As Double
    Dim endTN As Double
    Dim numList() As Double
    Dim numCount As Long
    Dim outputRange As Range
    Dim convertRng As String
    Dim output As Variant
    
    ClearRangeResult

    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    
    If lastRow < 3 Then
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
        If lastRow < 3 Then
            MsgBox "Done"
            Exit Sub
        End If
    End If
    
    ReDim output(1 To lastRow, 1 To 2)
    numCount = 0 ' Initialize array counter

    ' Loop through each TN range in column A
    For i = 3 To lastRow
        convertRng = CleanTN(ws.Cells(i, 1).value)
                
                
        If IsEmpty(ws.Cells(i, 2).value) And IsEmpty(ws.Cells(i, 3).value) Then
            startTN = Left(Trim(convertRng), 10)
            endTN = Left(Trim(convertRng), 6) & Right(Trim(ws.Cells(i, 1).value), 4)
            output(i - 2, 1) = startTN
            output(i - 2, 2) = endTN
        Else
            startTN = ws.Cells(i, 2).value
            endTN = ws.Cells(i, 3).value
        End If
        
        

        ' Store the TN numbers in an array
        For j = startTN To endTN
            numCount = numCount + 1
            ReDim Preserve numList(1 To numCount) ' Dynamically expand the array
            numList(numCount) = j
        Next j
    Next i

    ' Write data to Excel in bulk for better performance
    If numCount > 0 Then
        If IsEmpty(ws.Cells(3, 2).value) And IsEmpty(ws.Cells(3, 3).value) Then
            ws.Range("B3").Resize(lastRow, 2) = output
        End If
        Set outputRange = ws.Range("G3").Resize(numCount, 1)
        outputRange.value = Application.Transpose(numList)
    End If

    MsgBox "Done"
End Sub




Sub ClearRange()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).Row
    If lastRow > 2 Then
        Set rng = ws.Range("A3:G" & lastRow)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"
        
    End If
End Sub

Sub ClearRangeResult()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRowG As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRowA = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    lastRowG = ws.Cells(ws.Rows.count, 7).End(xlUp).Row

    If lastRowG > 2 Then
        Set rng = ws.Range("G3:G" & lastRowG)
        If lastRowA > 2 Then Set rng = ws.Range("B3:G" & lastRowG)
        
        rng.clear
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"
        
    End If
End Sub

Sub copyRange()

    Dim clipboardData As Object
    Dim lastRow As Long
    Dim CutInfo As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).Row
    For i = 3 To lastRow
        CutInfo = CutInfo & ws.Cells(i, 7).value & vbCrLf
    Next i
    
    
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText CutInfo
    clipboardData.PutInClipboard
End Sub

Function CleanTN(tn As String) As String
    Dim i As Integer
    Dim result As String
    Dim ch As String

    result = ""
    For i = 1 To Len(tn)
        ch = Mid(tn, i, 1)
        If ch Like "#" Then
            result = result & ch
        End If
    Next i

    CleanTN = result
End Function


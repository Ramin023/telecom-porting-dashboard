Attribute VB_Name = "Module4"
Sub ClearCompare() 'Compare two lists
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ClearCompareA
    ClearCompareB
    
End Sub

Sub ClearCompareA() 'Compare A lists
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow > 1 Then
        Set rng = ws.Range("A2:A" & lastRow)
        rng.ClearContents
    End If

End Sub

Sub ClearCompareB() 'Compare B lists
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 6).End(xlUp).Row
    If lastRow > 1 Then
        Set rng = ws.Range("F2:F" & lastRow)
        rng.ClearContents
    End If
    
End Sub

Sub copyYes()
    Dim clipboardData As Object
    Dim lastRow, startRow, list As Long
    Dim CutInfo As String
    
    Set ws = ActiveSheet
    starRow = 2
    list = 1
    lastRow = ws.Cells(ws.Rows.count, list).End(xlUp).Row
    
    For i = starRow To lastRow
        If ws.Cells(i, list + 1).value = "Yes" Then
            CutInfo = CutInfo & ws.Cells(i, list).value & vbCrLf
        End If
    Next i
    
    
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText CutInfo
    clipboardData.PutInClipboard
End Sub

Sub copyNoA()
    Dim clipboardData As Object
    Dim lastRow, startRow, list As Long
    Dim CutInfo As String
    
    Set ws = ActiveSheet
    starRow = 2
    list = 1
    lastRow = ws.Cells(ws.Rows.count, list).End(xlUp).Row
    
    For i = starRow To lastRow
        If ws.Cells(i, list + 1).value = "No" Then
            CutInfo = CutInfo & ws.Cells(i, list).value & vbCrLf
        End If
    Next i
    
    
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText CutInfo
    clipboardData.PutInClipboard
End Sub

Sub copyNoB()
    Dim clipboardData As Object
    Dim lastRow, startRow, list As Long
    Dim CutInfo As String
    
    Set ws = ActiveSheet
    starRow = 2
    list = 6
    lastRow = ws.Cells(ws.Rows.count, list).End(xlUp).Row
    
    For i = starRow To lastRow
        If ws.Cells(i, list + 1).value = "No" Then
            CutInfo = CutInfo & ws.Cells(i, list).value & vbCrLf
        End If
    Next i
    
    
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText CutInfo
    clipboardData.PutInClipboard
End Sub






Attribute VB_Name = "Module3"
Sub ListTNs() 'HiCap DID to List
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim npa As String
    Dim nxx As String
    Dim did As String
    Dim tn As String
    Dim numList() As Double
    Dim numCount As Long
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A (assuming data starts from row 4)
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    
    'exit if no paste
    If lastRow <= 3 Then
        MsgBox "Please copy DID detail in A2"
        Exit Sub
    End If
    
    For i = 1 To lastRow
        If ws.Cells(i, 1).value = "NPA" Then
            Exit For
        End If
    Next i
    
    ws.Cells(3, "G").value = "TNs"
    ws.Cells(3, "G").Font.Bold = True
    
    ' Loop through each row starting from row 4
    For i = i + 1 To lastRow
        ' Get the NPA, NXX, and DID
        If Not IsEmpty(ws.Cells(i, 1).value) Then
            npa = ws.Cells(i, 1).value
        End If
        If Not IsEmpty(ws.Cells(i, 2).value) Then
            nxx = ws.Cells(i, 2).value
        End If
        did = ws.Cells(i, 3).value
        
        
        ' Only proceed if there is a value in the DID column
        If IsNumeric(did) Then
            ' Construct the telephone number
            tn = npa & nxx & did
            
            ' Write the TN in the "TNs" column (column F)
            'ws.Cells(i, 7).value = tn
            numCount = numCount + 1
            ReDim Preserve numList(1 To numCount) ' Dynamically expand the array
            numList(numCount) = tn
        End If
    Next i
    
    If numCount > 0 Then
        Set outputRange = ws.Range("G4").Resize(numCount, 1)
        outputRange.value = Application.Transpose(numList)
    End If
    
    MsgBox "Done"
    
End Sub

Sub ClearList()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    If lastRow > 1 Then
        Set rng = ws.Range("A2:G" & lastRow)
        rng.clear

        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlLeft
        ws.Range("G:G").HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"

    End If
    
End Sub

Sub ClearListResult()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRow As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    If lastRow > 1 Then
        Set rng = ws.Range("G:G")
        rng.clear

        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"

    End If
    
End Sub

Sub copyList()
    Dim clipboardData As Object
    Dim lastRow As Long
    Dim CutInfo As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.count, 7).End(xlUp).Row
    For i = 4 To lastRow
        CutInfo = CutInfo & ws.Cells(i, 7).value & vbCrLf
    Next i
    
    
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText CutInfo
    clipboardData.PutInClipboard
End Sub




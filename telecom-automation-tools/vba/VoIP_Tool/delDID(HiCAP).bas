Attribute VB_Name = "Module5"
Sub delRange()
    Dim ws As Worksheet
    Dim lastRowA, lastRowC As Long
    Dim ClearRow As Long
    Dim newRow As Long
    Dim i As Long
    Dim j As Long
    Dim k As Double
    Dim rng As Range
    Dim delRange As Collection
    Dim startTN As Double
    Dim endTN As Double
    Dim startRange As Double
    Dim endRange As Double
    Dim aList As Collection
    Dim inRange As Boolean
    Dim newRange As Collection
    Dim count As Long
    Dim restOutput() As Variant
    Dim irest As Long

    
    ' Set the worksheet
    Set ws = ActiveSheet
    Set delRange = New Collection
    Set aList = New Collection
    Set newRange = New Collection
    count = 0
    
    ClearRowResult
          
    'last row of C
    lastRowC = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    'add A list TN to list
    lastRowA = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    If lastRowA = 2 Or lastRowC = 1 Then
        MsgBox "Done"
        Exit Sub
    End If
    
    
    For i = lastRowC To 2 Step -1
        If ws.Cells(i, 3).value = "Edit" Then
            ClearRow = i + 1
            Exit For
        End If
    Next i
    
    
    If Not (ws.Cells(ClearRow, 4).value = "" And ws.Cells(ClearRow, 5).value = "") Then
        ws.Cells(ClearRow, 9).value = ws.Cells(ClearRow, 4).value
        ws.Cells(ClearRow, 10).value = ws.Cells(ClearRow, 5).value
        ws.Cells(ClearRow, 4).value = ""
        ws.Cells(ClearRow, 5).value = ""
    End If
    
    
    
    If ClearRow > 1 And ClearRow <= lastRowC Then
        Set rng = ws.Range("C" & ClearRow + 1 & ":k" & lastRowC)
        rng.ClearContents
    End If
    
    
    
    
    ws.Range("A2:A" & lastRowA).Sort Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes

    
    For i = 3 To lastRowA
        aList.Add ws.Cells(i, 1).value
    Next i

    
    'filter del range
    For j = 4 To ClearRow - 1
        startTN = ws.Cells(j, 4).value & ws.Cells(j, 5).value & ws.Cells(j, 6).value
        endTN = ws.Cells(j, 4).value & ws.Cells(j, 5).value & ws.Cells(j, 8).value
        inRange = False
        
        startRange = 0
        
        For k = startTN To endTN
            ' Check if current value is in the removed collection
            If Not IsInCollection(k, aList) Then
                If startRange = 0 Then startRange = k
                endRange = k
            Else
                RemoveFromCollection aList, k
                inRange = True
                
                If startRange > 0 Then
                    newRange.Add startRange & " " & endRange
                    startRange = 0
                End If
                
                If aList.count = 0 Then Exit For
            End If
        Next k
        
        If startRange > 0 And inRange Then
            newRange.Add startRange & " " & endRange
            startRange = 0
        End If
        
        If inRange Then delRange.Add startTN & " " & endTN
        If aList.count = 0 Then
            If Not k = endTN Then
                newRange.Add k + 1 & " " & endTN
            End If
            Exit For
        End If
    Next j
    
    newRow = ClearRow + 2
    ' print out
    If delRange.count > 0 Then
        'print out del Range
        ws.Cells(newRow, 3).value = "Delete Range"
        ws.Cells(newRow, 3).Font.Bold = True
        newRow = newRow + 1
        
        Dim delOutput() As Variant
        Dim iDel As Long
        
        ReDim delOutput(1 To delRange.count, 1 To 8)
        iDel = 1
        
        For Each item In delRange
            delOutput(iDel, 1) = Left(item, 3)
            delOutput(iDel, 2) = Mid(item, 4, 3)
            delOutput(iDel, 3) = Mid(item, 7, 4)
            delOutput(iDel, 4) = "->"
            delOutput(iDel, 5) = Right(item, 4)
            delOutput(iDel, 6) = "="
            delOutput(iDel, 7) = Right(item, 4) - Mid(item, 7, 4) + 1
            count = count + delOutput(iDel, 7)
            iDel = iDel + 1
        Next item
        
        ws.Range(ws.Cells(newRow, 4), ws.Cells(newRow + delRange.count - 1, 10)).value = delOutput
        newRow = newRow + delRange.count
        ws.Cells(newRow, 9).value = "Remove: "
        ws.Cells(newRow, 10).value = count
        count = 0
        
        If newRange.count > 0 Then
            newRow = newRow + 2
            ws.Cells(newRow, 3).value = "Rebuilt Range"
            ws.Cells(newRow, 3).Font.Bold = True
             newRow = newRow + 1
            
            Dim newOutput() As Variant
            Dim inew As Long
            ReDim newOutput(1 To newRange.count, 1 To 8)
            inew = 1
           
            'print out new range
            
            For Each item In newRange
                newOutput(inew, 1) = Left(item, 3)
                newOutput(inew, 2) = Mid(item, 4, 3)
                newOutput(inew, 3) = Mid(item, 7, 4)
                newOutput(inew, 4) = "->"
                newOutput(inew, 5) = Right(item, 4)
                newOutput(inew, 6) = "="
                newOutput(inew, 7) = Right(item, 4) - Mid(item, 7, 4) + 1
                count = count + newOutput(inew, 7)
                inew = inew + 1
            Next item
            
            ws.Range(ws.Cells(newRow, 4), ws.Cells(newRow + newRange.count - 1, 10)).value = newOutput
            newRow = newRow + newRange.count
            ws.Cells(newRow, 9).value = "Re-add: "
            ws.Cells(newRow, 10).value = count
            count = 0
        End If
        
        'print out remian TNs
        If aList.count > 0 Then
            newRow = newRow + 2
            ws.Cells(newRow, 3).value = "TNs are't in these Ranges"
            ws.Cells(newRow, 3).Font.Bold = True
            newRow = newRow + 1

            ReDim restOutput(1 To aList.count, 1 To 8)
            irest = 1
            
            For Each item In aList
                restOutput(irest, 1) = item
                irest = irest + 1
            Next item
            
            ws.Range(ws.Cells(newRow, 3), ws.Cells(newRow + aList.count - 1, 3)).value = restOutput
            newRow = newRow + aList.count
        End If
    Else 'print out remian TNs if no range is del
        ws.Cells(newRow, 3).value = "TNs are't in these Ranges"
        ws.Cells(newRow, 3).Font.Bold = True
        newRow = newRow + 1
        
        ReDim restOutput(1 To aList.count, 1 To 8)
        irest = 1
        
        For Each item In aList
            restOutput(irest, 1) = item
            irest = irest + 1
        Next item
        
        ws.Range(ws.Cells(newRow, 3), ws.Cells(newRow + aList.count - 1, 3)).value = restOutput
        newRow = newRow + aList.count
    End If
        
    ws.Cells(newRow, 3).Select
       
    MsgBox "Done"
End Sub
   
Function RemoveFromCollection(coll As Collection, value As Double)
    Dim i As Long
    On Error Resume Next
    For i = 1 To coll.count
        If coll(i) = value Then
            coll.Remove i
            Exit For
        End If
    Next i
    On Error GoTo 0
End Function


Function IsInCollection(value As Double, coll As Collection) As Boolean
    Dim item As Variant
    On Error Resume Next
    IsInCollection = False
    For Each item In coll
        If item = value Then
            IsInCollection = True
            Exit Function
        End If
    Next item
    On Error GoTo 0
End Function

    
Sub ClearRow()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRowA, lastRowC, lastRowJ As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    
    lastRowA = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRowA > 2 Then
        Set rng = ws.Range("A3:A" & lastRowA)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlCenter
        rng.NumberFormat = "@"
    End If
    
    lastRowC = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    lastRowJ = ws.Cells(ws.Rows.count, 10).End(xlUp).Row
    
    If lastRowC > lastRowJ Then
        Set rng = ws.Range("C2:J" & lastRowC)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlLeft
        rng.NumberFormat = "@"
    Else
        Set rng = ws.Range("C2:J" & lastRowJ)
        rng.clear
        
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlLeft
        rng.NumberFormat = "@"
    End If

    
End Sub

Sub ClearRowResult()
    Dim ws As Worksheet
    Dim rng As Range
    Dim i, startRow, lastRowC, lastRowJ As Long
    
    ' Set the worksheet object
    Set ws = ActiveSheet
    startRow = 3
    lastRowC = ws.Cells(ws.Rows.count, 3).End(xlUp).Row
    lastRowJ = ws.Cells(ws.Rows.count, 10).End(xlUp).Row
    
    If lastRowC > 2 And lastRowJ > 2 Then
        For i = startRow To lastRowJ
            If ws.Cells(i, 4).value = "" Then
                startRow = i
                Exit For
            End If
        Next i
        
        If startRow = 3 Then startRow = lastRowJ + 1

        
        If lastRowC > lastRowJ Then
            Set rng = ws.Range("C" & startRow + 2 & ":J" & lastRowC)
            rng.clear
            
            rng.Font.Name = "Aptos Narrow"
            rng.HorizontalAlignment = xlLeft
            rng.NumberFormat = "@"
        Else
            Set rng = ws.Range("C" & startRow + 2 & ":J" & lastRowJ)
            rng.clear
            
            rng.Font.Name = "Aptos Narrow"
            rng.HorizontalAlignment = xlLeft
            rng.NumberFormat = "@"
        End If
    End If
        
End Sub


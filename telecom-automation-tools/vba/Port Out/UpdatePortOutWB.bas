Attribute VB_Name = "Module1"
Sub UpdatePortOutWB()
    Dim wb As Workbook
    Dim wsTarget As Worksheet, wsSource As Worksheet
    Dim lastRowTarget As Long, lastRowSource As Long
    Dim endRow As Long, startRow As Long
    Dim i As Long
    Dim adddate As String
    Dim sourceData As Variant, targetData As Variant
    Dim newData() As Variant
    Dim newCount As Long
    Dim newRange As Range
    Dim cell As Range
    Dim dict As Object
    Dim dataArr As Variant
    Dim Key As Variant
    Dim arr() As String
    Dim delRange As Range
    Dim dupCount As Long

    ' Set worksheets
    Set wb = ActiveWorkbook
    Set wsTarget = wb.Sheets("VoIP")    ' Target sheet
    Set wsSource = wb.Sheets("PortOut") ' Source sheet
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Find last row
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row
    
    ' Store the current date
    adddate = Format(Date, "mm/dd")
    endRow = lastRowTarget
    startRow = endRow

    ' Load data into arrays
    sourceData = wsSource.Range("A2:N" & lastRowSource).Value

    ' Preallocate array
    ReDim newData(1 To lastRowSource, 1 To 7)
    newCount = 0

    ' Sort by column F
    With wsSource.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsSource.Range("F:F"), Order:=xlAscending
        .SetRange wsSource.Range("A1").CurrentRegion
        .Header = xlYes
        .Apply
    End With

    ' Loop through source data
    For i = 1 To lastRowSource - 1
        If Not InStr(1, sourceData(i, 4), "RCF", vbTextCompare) > 0 Then
            newCount = newCount + 1
            newData(newCount, 1) = sourceData(i, 1)
            newData(newCount, 2) = CStr(sourceData(i, 2))
            newData(newCount, 3) = sourceData(i, 5)
            newData(newCount, 7) = sourceData(i, 14)

            ' Handle status
            If Not IsEmpty(sourceData(i, 11)) Then
                If newData(newCount, 1) = "Bandwidth" Then
                    newData(newCount, 4) = "Completed"
                Else
                    newData(newCount, 4) = "Confirmed"
                End If
                newData(newCount, 5) = Format(sourceData(i, 11), "mm/dd/yyyy")
            Else
                newData(newCount, 4) = "Pending " & adddate
            End If
        End If
    Next i

    ' Write new data in bulk
    If newCount > 0 Then
        Set newRange = wsTarget.Range("A" & lastRowTarget + 1).Resize(newCount, 7)
        newRange.Value = newData

        ' Change "Completed" cells to blue
        For Each cell In newRange.Columns(3).Cells
            If cell.Value = "Completed" Then
                cell.Interior.Color = RGB(0, 176, 240)
            End If
        Next cell
    End If

    ' Mark and remove duplicates (DUP)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "B").End(xlUp).Row
    If lastRowTarget < 2 Then Exit Sub

    ' Load data
    dataArr = wsTarget.Range("B2:B" & lastRowTarget).Value
    Set dict = CreateObject("Scripting.Dictionary")
    dupCount = 0

    ' Count occurrences
    For i = 1 To UBound(dataArr)
        If dict.exists(dataArr(i, 1)) Then
            dict(dataArr(i, 1)) = dict(dataArr(i, 1)) & "," & i
        Else
            dict.Add dataArr(i, 1), i
        End If
    Next i

    ' Mark "DUP"
    For Each Key In dict.keys
        arr = Split(dict(Key), ",")
        If UBound(arr) >= 1 Then
            For i = 0 To UBound(arr) - 1
                wsTarget.Cells(arr(i) + 1, 5).Value = "DUP"
                dupCount = dupCount + 1
            Next i
        End If
    Next Key

    ' Delete "DUP" rows
    With wsTarget.Range("A1:F" & lastRowTarget)
        .AutoFilter Field:=5, Criteria1:="DUP"
        On Error Resume Next
        Set delRange = wsTarget.Range("A2:A" & lastRowTarget).SpecialCells(xlCellTypeVisible).EntireRow
        On Error GoTo 0
        If Not delRange Is Nothing Then delRange.Delete
        .AutoFilter
    End With
    Debug.Print newCount
    
    For i = endRow - dupCount To endRow
        If wsTarget.Cells(i, 1).Value = "Bandwidth" Then
            wsTarget.Cells(i, 4).Interior.Color = RGB(0, 176, 240)
        End If
    Next i

    ' Restore screen updates
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    wsTarget.Cells(endRow - dupCount, 6).Select

    ' Show summary message
    MsgBox "New data added: " & newCount - dupCount & vbNewLine & _
           "Updated data: " & dupCount, vbInformation
End Sub




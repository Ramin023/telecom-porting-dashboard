Attribute VB_Name = "Module7"
Sub ParseAddress(mode As String)
    Dim parts() As String
    Dim i As Long, j As Long, lastRow As Long
    Dim streetNumber As String, streetName As String, streetType As String
    Dim locationInfo As String, city As String
    Dim state As String, zipCode As String
    Dim perD As String, postD As String
    Dim locStart As Long, cityStart As Long, stEnd As Long, locEnd As Long, stStart
    Dim ws As Worksheet
    Dim output() As Variant
    
    Set ws = ActiveSheet

    Dim locKeywords As Variant
    locKeywords = Array("apt", "apartment", "suite", "suit", "ste", "unit", "bldg", "building", _
                        "fl", "floor", "room", "rm", "dept", "department", "lot", "trlr", "trailer", _
                        "hangar", "pier", "slip", "stop", "space", "box", "po", "p.o.")

    Dim streetKeywords As Variant
    streetKeywords = Array("st", "street", "ave", "avenue", "av", "rd", "road", "blvd", _
                           "dr", "drive", "ln", "lane", "ct", "court", "pl", "place", "cir", "circle", _
                           "ter", "terrace", "way", "freeway", "trl", "trail", "pkwy", "parkway", _
                           "aly", "alley", "loop", "sq", "square", "cres", "crescent", "pt", "point", _
                           "drwy", "driveway", "plz")
                           
    Dim directionKeywords As Variant
    directionKeywords = Array("s", "e", "w", "n", "sw", "se", "nw", "ne")

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
    
    Dim outCols As Integer
    Select Case mode
        Case "type1"
            outCols = 10
        Case "type2"
            outCols = 7
        Case "type3"
            outCols = 6
    End Select
    
    ReDim output(1 To lastRow, 1 To outCols)
    
    Select Case mode
        Case "type1"
            output(1, 1) = "Street Number"
            output(1, 2) = "Street Pre Direction"
            output(1, 3) = "Street Name"
            output(1, 4) = "Street Type"
            output(1, 5) = "Street Post Direction"
            output(1, 6) = "LOC"
            output(1, 7) = "City"
            output(1, 8) = "State"
            output(1, 9) = "ZIP"
        Case "type2"
            output(1, 1) = "Street Number"
            output(1, 2) = "Street Name"
            output(1, 3) = "LOC"
            output(1, 4) = "City"
            output(1, 5) = "State"
            output(1, 6) = "ZIP"
        Case "type3"
            output(1, 1) = "Address 1"
            output(1, 2) = "Address 2"
            output(1, 3) = "City"
            output(1, 4) = "State"
            output(1, 5) = "ZIP"
    End Select

    For i = 3 To lastRow
        If Trim(ws.Cells(i, 1).value) <> "" Then
            parts = Split(Replace(Replace(Trim(ws.Cells(i, 1).value), ",", ""), ".", ""), " ")
            If UBound(parts) < 4 Then
                ws.Cells(i, 2).value = "Invalid"
                GoTo SkipRow
            End If
            If IsNumeric(parts(UBound(parts))) Then
                zipCode = Left(parts(UBound(parts)), 5)
                state = parts(UBound(parts) - 1)
                cityStart = UBound(parts) - 2
                
            Else
                zipCode = parts(UBound(parts))
                If Len(zipCode) <= 3 Then
                    zipCode = parts(UBound(parts) - 1) & zipCode
                    state = parts(UBound(parts) - 2)
                    cityStart = UBound(parts) - 3
                Else
                    state = parts(UBound(parts) - 1)
                    cityStart = UBound(parts) - 2
                End If
            End If
            
            streetNumber = parts(0)

            locStart = -1
            locEnd = -1
            stEnd = -1
            stStart = 1
            
            streetName = ""
            city = ""
            locationInfo = ""
            streetType = ""
            postD = ""
            perD = ""
            
            If IsInArray(LCase(parts(1)), directionKeywords) Then
                perD = parts(1)
                stStart = 2
            End If
            
            For j = stStart To cityStart
                If IsInArray(LCase(parts(j)), locKeywords) Then
                    If locStart > 0 Then locEnd = j + 1 Else locStart = j
                End If
            Next j
     
            If locEnd < 0 Then locEnd = locStart + 1
            If locEnd >= cityStart Then locEnd = cityStart - 1
            
            For j = stStart To cityStart - 1
                If IsInArray(LCase(parts(j)), streetKeywords) Then
                    stEnd = j
                End If
            Next j
            
            If locStart < 0 Then 'No LOC in address
                
                If stEnd > 0 Then
                    For j = stStart To stEnd - 1: streetName = streetName & parts(j) & " ": Next j
                    streetType = parts(stEnd)
                    If IsInArray(LCase(parts(stEnd + 1)), directionKeywords) Then
                        postD = parts(stEnd + 1)
                        stEnd = stEnd + 1
                    End If
                    For j = stEnd + 1 To cityStart: city = city & parts(j) & " ": Next j
                Else
                    For j = stStart To cityStart - 1: streetName = streetName & parts(j) & " ": Next j
                    city = parts(cityStart)
                End If
            Else ' LOC in address
                stEnd = locStart - 1
                
                If IsInArray(LCase(parts(stEnd)), streetKeywords) Then
                    For j = stStart To stEnd - 1: streetName = streetName & parts(j) & " ": Next j
                    streetType = parts(stEnd)
                ElseIf IsInArray(LCase(parts(stEnd - 1)), streetKeywords) Then
                    For j = stStart To stEnd - 2: streetName = streetName & parts(j) & " ": Next j
                    streetType = parts(stEnd - 1)
                    postD = parts(stEnd)
                Else
                    For j = stStart To locStart - 1: streetName = streetName & parts(j) & " ": Next j
                End If
                
                For j = locStart To locEnd: locationInfo = locationInfo & parts(j) & " ": Next j
                For j = locEnd + 1 To cityStart: city = city & parts(j) & " ": Next j
            End If

            streetName = Trim(streetName)
            locationInfo = Trim(locationInfo)
            city = Trim(city)

            Select Case mode
                Case "type1"
                    output(i - 1, 1) = streetNumber
                    output(i - 1, 2) = perD
                    output(i - 1, 3) = streetName
                    output(i - 1, 4) = streetType
                    output(i - 1, 5) = postD
                    output(i - 1, 6) = locationInfo
                    output(i - 1, 7) = city
                    output(i - 1, 8) = state
                    output(i - 1, 9) = zipCode
                Case "type2"
                    output(i - 1, 1) = streetNumber
                    output(i - 1, 2) = Trim(perD & " " & streetName & " " & streetType & " " & postD)
                    output(i - 1, 3) = locationInfo
                    output(i - 1, 4) = city
                    output(i - 1, 5) = state
                    output(i - 1, 6) = zipCode
                Case "type3"
                    output(i - 1, 1) = Replace(Trim(streetNumber & " " & perD & " " & streetName & " " & streetType & " " & postD), "  ", " ")
                    output(i - 1, 2) = locationInfo
                    output(i - 1, 3) = city
                    output(i - 1, 4) = state
                    output(i - 1, 5) = zipCode
            End Select
        End If
SkipRow:
    Next i
    
    ClearAddressResult
    ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, outCols)).value = output
End Sub

Function IsInArray(val As String, arr As Variant) As Boolean
    Dim x
    For Each x In arr
        If val = x Then IsInArray = True: Exit Function
    Next x
    IsInArray = False
End Function

Sub Tyep1()
    Call ParseAddress("type1")
End Sub

Sub Tyep2()
    Call ParseAddress("type2")
End Sub

Sub Tyep3()
    Call ParseAddress("type3")
End Sub

Sub ClearAddressResult()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRowB, startRow As Long
    
    startRow = 2
    ' Set the worksheet object
    Set ws = ActiveSheet
    
    lastRowB = ws.Cells(ws.Rows.count, 2).End(xlUp).Row
    
    If lastRowB > 2 Then
        Set rng = ws.Range("B" & startRow & ":K" & lastRowB)
        rng.clear
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlLeft
        rng.NumberFormat = "@"
    End If
End Sub

Sub ClearAddressAll()
    Dim ws As Worksheet
    Dim rng As Range
    Dim lastRowA, startRow As Long
    
    startRow = 2
    ' Set the worksheet object
    Set ws = ActiveSheet
    
    lastRowA = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    ClearAddressResult
    
    If lastRowA > 2 Then
        Set rng = ws.Range("A" & startRow + 1 & ":A" & lastRowA)
        rng.clear
        rng.Font.Name = "Aptos Narrow"
        rng.HorizontalAlignment = xlLeft
        rng.NumberFormat = "@"
    End If
    
    
End Sub


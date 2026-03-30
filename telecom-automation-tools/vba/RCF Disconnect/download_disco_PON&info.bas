Attribute VB_Name = "Module3"
Sub run()
    Dim http As Object
    Dim html As Object
    Dim htmlText As String
    Dim table As Object
    Dim rows As Object
    Dim row As Object
    Dim cells As Object
    Dim i As Long, n As Long
    Dim manager As String
    Dim pon As String, dueDate As String, formattedDate As String, tn As String
    Dim today As Date
    
    
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Find the last row in column A
    lastRow = ws.cells(ws.rows.Count, "B").End(xlUp).row
    
    today = Date
    manager = "ZACHARY TANTILLO"
    n = 1
    ReDim output(1 To n, 1 To 10)
    
    Dim urlInitiated As String
    urlInitiated = "http://tickets.granitenet.com/ReportsNet/ReportsCornerstone/OpenMacsReport.aspx?Login=&SystemID=&ReportNumber=36&VAL=25&Sort= scrolling =auto"

    Dim urlError As String
    urlError = "http://tickets.granitenet.com/ReportsNet/ReportsCornerstone/OpenMacsReport.aspx?Login=&SystemID=&ReportNumber=36&VAL=5&Sort= scrolling =auto"
    Dim urlReject As String
    urlReject = "http://tickets.granitenet.com/ReportsNet/ReportsCornerstone/OpenMacsReport.aspx?Login=&SystemID=&ReportNumber=36&VAL=2&Sort= scrolling =auto"


    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", urlReject, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send
    htmlText = http.responseText
    

    Set html = CreateObject("htmlfile")
    html.body.innerHTML = htmlText
    

    Set table = html.getElementById("dgReport")
    If table Is Nothing Then
        MsgBox "找不到表格 dgReport"
        Exit Sub
    End If
    
    Set rows = table.getElementsByTagName("tr")
    

    For i = 1 To rows.Length - 1
        Set row = rows.Item(i)
        Set cells = row.getElementsByTagName("td")
        
        
        If cells.Length >= 8 Then
            
            If Trim(cells.Item(2).innerText) = manager Then
                
                pon = Trim(cells.Item(3).innerText)
                dueDate = Trim(cells.Item(8).innerText)
                formattedDate = Format(DateSerial(Left(dueDate, 4), Mid(dueDate, 5, 2), Right(dueDate, 2)), "mm/dd/yyyy")
                Debug.Print "PON#: " & pon & " | Due Date: " & formattedDate
                

                Dim urlCSdisc As String
                'urlCSdisc = "http://cornerstone.granitenet.com/live/local/lc_disc.asp?PON_TAG_LOCATE=&ReqType=E&action=D&ILEC=QWEST&pon=GRTN806829691&ver=00&OP_ID=&PON_EX_ACTION=Disconnect&GetOldVER=t"
                'http.Open "GET", urlCSdisc, False
                'http.setRequestHeader "User-Agent", "Mozilla/5.0"
                'http.send
                'htmlText1 = http.responseText
                
                tn = ""
                
                For j = 1 To 128
                    Set regex = CreateObject("VBScript.RegExp")
                    regex.Pattern = "name\s*=\s*[""']db_" & j & "_(DISCNBR|TCTOPRI|TCPER)[""']\s+[^>]*value\s*=\s*[""']([^""']+)[""']"
                    regex.IgnoreCase = True
                    regex.Global = True
                
                    Set matches = regex.Execute(htmlText1)
                
                    If matches.Count > 0 Then
                        tn = ""
                        ttn = ""
                        tday = 0
                    

                        For Each Match In matches
                            Select Case UCase(Match.SubMatches(0))
                                Case "DISCNBR"
                                    tn = Match.SubMatches(1)
                                Case "TCTOPRI"
                                    ttn = Match.SubMatches(1)
                                Case "TCPER"
                                    tday = Match.SubMatches(1)
                            End Select
                        Next
                        
                        output(n, 1) = "OM"
                        output(n, 2) = pon
                        output(n, 3) = tn
                        output(n, 5) = today
                        If today < formattedDate Then
                            output(n, 6) = today
                        Else
                            output(n, 6) = formattedDate
                        End If
                        

                        
                        If tday > 0 Then
                            output(n, 8) = "Yes"
                            output(n, 9) = CDate(formattedDate) + tday
                            output(n, 10) = ttn
                        Else
                            output(n, 8) = "No"
                            output(n, 9) = "N/A"
                            output(n, 10) = "N/A"
                        End If
                    Else
                        Exit For
                    End If
                    n = n + 1
                Next j
            End If
        End If
    Next i
    
    Debug.Print tn
    Debug.Print ttn
    Debug.Print CDate(formattedDate) + tday
    ws.Range(ws.cells(lastRow + 1, "B"), ws.cells(lastRow + n - 1, "K")).Value = output
    'Debug.Print htmlText1

End Sub



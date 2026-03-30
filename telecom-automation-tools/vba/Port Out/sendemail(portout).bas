Attribute VB_Name = "Module2"
Function Replymail(ByVal tracker As String, ByVal tnsTable As String) As Object
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olRootFolder As Outlook.Folder
    Dim olFolder As Outlook.Folder
    Dim olItems As Outlook.Items
    Dim filteredItems As Outlook.Items
    Dim filter As String
    Dim newestItem As MailItem
    Dim bodyHTML As String
    Dim olReply As Outlook.MailItem

    'tracker = "-7960827107406263163" ' Replace with dynamic value if needed

    Set olApp = Outlook.Application
    Set olNs = olApp.GetNamespace("MAPI")
    
    ' Reference the shared mailbox root
    Set olRootFolder = olNs.GetDefaultFolder(olFolderInbox).Parent ' Adjust this if the name is different in your Outlook
    
    ' Get the "Port Out" subfolder
    Set olFolder = olRootFolder.Folders("Port Out")

    ' Get and sort the items by ReceivedTime (newest first)
    Set olItems = olFolder.Items
    olItems.Sort "[ReceivedTime]", True

    ' DASL filter for subject containing tracker
    filter = "@SQL=" & _
         """urn:schemas:mailheader:subject"" LIKE '%" & tracker & "%'"


    ' Apply filter
    Set filteredItems = olItems.Restrict(filter)

    
    bodyHTML = "<html><body style=""font-family: 'Times New Roman', Times, serif; font-size: 11pt;"">" & _
                "Hello, <br><br>" & _
                "The below lines have ported away: <br><br>" & _
                tnsTable & _
                "<br/>" & _
                "Please confirm with the customer whether this was an intentional port-out.<br>" & _
                "<ul>" & _
                "<li>If <strong>intentional</strong>, open a <strong>Disconnect Ticket</strong> to remove the number(s) from the switch and DID details.</li>" & _
                "<li>If <strong>unintentional</strong>, open a <strong>Reinstate Ticket</strong> so we can attempt to reclaim the line.</li>" & _
                "</ul>" & _
                "When submitting the ticket, note that this is due to a <strong>Port Out</strong>. If <strong>Early Termination Fees (ETFs)</strong> apply, include them in the Disconnect Ticket.<br>" & _
                "Please refer to the <strong>DID Disconnect Ticket Template</strong> available in <strong>Shelf</strong> for guidance.<br>" & _
                "<p>Thank you.</p>" & _
                "</body></html>"
                

    If filteredItems.Count > 0 Then
        Dim lastEmail As Object
        Set lastEmail = filteredItems.Item(1)
        ' Reply to the last email

        Set olReply = lastEmail.ReplyAll
        ' Modify the replyMail object as needed (e.g., adding text)
        
        'remove Kevin Russell
        Dim i As Integer
        For i = olReply.Recipients.Count To 1 Step -1
            If olReply.Recipients(i).Name = "Kevin Russell" Then
                olReply.Recipients(i).Delete
            End If
        Next i
        
        olReply.HTMLBody = bodyHTML & olReply.HTMLBody
        ' Send the reply
        olReply.Display
    Else
        MsgBox "No matching emails found for tracker: " & tracker
    End If
End Function


Sub fileterAndSend()
    Dim wsTarget As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim trackerNumber, tnsTable, account, tns As String
    
    ' Set the worksheet
    Set wsTarget = ThisWorkbook.Sheets("VoIP") ' Source worksheet
    
    wsTarget.AutoFilterMode = False
    
    ' Find the last row in column A
    lastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    
    ' Apply filter to Column C for cells with "Completed"
    wsTarget.Range("A1:G1").AutoFilter Field:=4, Criteria1:="Completed", Operator:=xlFilterValues
    
    ' Apply filter to Column E for empty cells
    wsTarget.Range("A1:G1").AutoFilter Field:=6, Criteria1:="", Operator:=xlFilterValues

    ' Sort the filtered range by column F
    wsTarget.Range("A1:G" & lastRow).Sort Key1:=wsTarget.Range("G1"), Order1:=xlAscending, Header:=xlYes
    
    
    ' Set rng to the visible cells in column A
    On Error Resume Next
    Set rng = wsTarget.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If rng Is Nothing Then
        wsTarget.AutoFilterMode = False
        wsTarget.Range("A1:G1").AutoFilter
        wsTarget.Range("A1:G" & lastRow).Sort Key1:=wsTarget.Range("E1"), Order1:=xlAscending, Header:=xlYes
        wsTarget.Range("A1:G" & lastRow).Sort Key1:=wsTarget.Range("F1"), Order1:=xlAscending, Header:=xlYes
        
        MsgBox "No more Email sent", vbInformation
        Exit Sub ' Exit the subroutine or add further actions here
    End If
    
    trackerNumber = rng.Rows(1).Cells.Offset(0, 5).Value
    'tnsTable = ""
    tnsTable = "<table border='0' cellspacing='5' cellpadding='0' style='font-family:Times New Roman; font-size:11pt;'>"
    'tnsTable = tnsTable & "<tr style='background-color: #f2f2f2;'><td style='width:100px; padding:8px;'><b>Account</b></td><td><b>TN</b></td></tr>"
    
    account = ""
    
    ' Loop through each cell in rng
    For Each cell In rng
        cell.Offset(0, 4).Value = cell.Offset(0, 3).Value
        wsTarget.Range(cell.Offset(0, -1), cell.Offset(0, 5)).Interior.Color = RGB(0, 176, 240)

        If trackerNumber <> cell.Offset(0, 5).Value Then
            Replymail trackerNumber, tnsTable & "</table>"
            tnsTable = "<table border='0' cellspacing='5' cellpadding='0' style='font-family:Times New Roman; font-size:11pt;'>"
            'tnsTable = tnsTable & "<tr style='background-color: #f2f2f2;'><td style='width:100px; padding:8px;'><b>Account</b></td><td><b>TN</b></td></tr>"
            account = ""
        End If
        
        If cell.Offset(0, 1).Value = "N/A" Then
            tnsTable = tnsTable & cell.Offset(0, 0).Value & "<br>"
        ElseIf account <> cell.Offset(0, 1).Value Then
            account = cell.Offset(0, 1).Value
            tnsTable = tnsTable & "<tr><td style='width:100px;'>" & account & ":</td><td>" & cell.Offset(0, 0).Value & "</td></tr>"
        Else
            tnsTable = tnsTable & "<tr><td style='width:100px;'></td><td>" & cell.Offset(0, 0).Value & "</td></tr>"
        End If
        

        trackerNumber = cell.Offset(0, 5).Value
    Next cell
        
    'send last email
    Replymail trackerNumber, tnsTable & "</table>"
    
    ' Clear the filter in column E
    wsTarget.AutoFilterMode = False
    wsTarget.Range("A1:G1").AutoFilter
    
    wsTarget.Range("A1:G" & lastRow).Sort Key1:=wsTarget.Range("E1"), Order1:=xlAscending, Header:=xlYes
    wsTarget.Range("A1:G" & lastRow).Sort Key1:=wsTarget.Range("F1"), Order1:=xlAscending, Header:=xlYes
    
    MsgBox "Done"
End Sub





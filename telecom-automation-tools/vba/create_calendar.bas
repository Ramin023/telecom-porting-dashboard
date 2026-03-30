Attribute VB_Name = "Module1"
Sub createWONMetting()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olRecip As Object
    Dim olCalFolder As Object
    Dim olApt As Object
    Dim wo As String
    Dim CutInfo As String
    

    Set selectedRow = Selection.EntireRow

    Dim completeDate As String
    completeDate = Format(selectedRow.Cells(9).Value, "mm/dd")
    
    wo = selectedRow.Cells(3).Value
    'If selectedRow.Cells(2).Value <> "" Then
        'wo = selectedRow.Cells(2).Value & "-" & selectedRow.Cells(3).Value
    'End If
    
    
    Dim Title As String
    'Title = "ZTP " & selectedRow.Cells(1).Value & " LNP|" & wo & _
            '"|" & selectedRow.Cells(4).Value & _
            '"|" & selectedRow.Cells(5).Value & _
            '"|" & selectedRow.Cells(6).Value
            
    Title = "LNP | " & selectedRow.Cells(5).Value & " | " & _
            selectedRow.Cells(6).Value & " |" & _
            selectedRow.Cells(4).Value & "|" & _
            wo
            

    ' Modify Body to make "FOC Received" bold
    Dim emailBody As String

    emailBody = "**FOC Received**" & vbCrLf & vbCrLf & _
                "**FOC Pending**" & vbCrLf & vbCrLf & _
                "Order#: " & selectedRow.Cells(10).Value & vbCrLf & _
                "TN: " & selectedRow.Cells(12).Value & vbCrLf & _
                "Request DD: " & selectedRow.Cells(9).Value & vbCrLf & vbCrLf & _
                "*** Please email VoIP Triggers on DD to trigger port ***" & vbCrLf & vbCrLf & _
                "Please inform the customer that a VoIP SERVICE ORDER CHARGE of $45 or " & _
                "PORT RESCHEDULE UNDER 24 HOURS CHARGE of $200 (if requested within 24 hours) will be applied " & _
                "to the account once the order is FOC and the customer requests to delay the due date."
                

    ' Create a new instance of Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    If olApp Is Nothing Then
        ' Outlook is not running, so create a new instance
        Set olApp = CreateObject("Outlook.Application")
    End If
    
    ' Get the Outlook Namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Add the shared calendar to the namespace
    Set olRecip = olNamespace.CreateRecipient("NetworkActivationsCalendar")
    olRecip.Resolve

    If olRecip.Resolved Then
        ' Access the shared calendar folder
        Set olCalFolder = olNamespace.GetSharedDefaultFolder(olRecip, 9) ' 9 represents olFolderCalendar
        ' Create a new AppointmentItem in the shared calendar
        Set olApt = olCalFolder.Items.Add(1) ' 1 represents olAppointmentItem
        
        ' Set appointment properties
        With olApt
            .Subject = Title
            .Start = CDate(selectedRow.Cells(9).Value & " 15:00:00") ' Set start time
            .End = CDate(selectedRow.Cells(9).Value & " 15:30:00")   ' Set end time
            .Location = "Trigeronly"
            .body = emailBody
            .ReminderSet = True
            .ReminderMinutesBeforeStart = 15 ' Set reminder time in minutes
            .Recipients.Add "mlei@granitenet.com" ' Add a recipient
            .Recipients.ResolveAll ' Resolve recipients' names
            .Display ' Save the appointment
        End With
    End If
    
    ' Clean up
    Set olApt = Nothing
    Set olCalFolder = Nothing
    Set olRecip = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub





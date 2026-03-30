Attribute VB_Name = "Module5"
Sub searchWONMeeting()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olCalFolder As Object
    Dim foundItems As Object
    Dim item As Object
    Dim ticket, acc As String
    Dim searchFilter As String
    Dim clipboardData As Object
    Dim selectedRow As Range
    Dim newStartDate As Date
    Dim newEndDate As Date
    Dim olRecip As Object

    ' Get the selected row and retrieve the PWO
    Set selectedRow = Selection.EntireRow
    'If selectedRow.Cells(2).Value <> "" Then
        'pwo = selectedRow.Cells(2).Value & "-" & selectedRow.Cells(3).Value
    'Else
        'pwo = selectedRow.Cells(3).Value & "|" & selectedRow.Cells(4).Value
    'End If

    ticket = selectedRow.Cells(3).Value
    acc = selectedRow.Cells(4).Value
    
    If Left(ticket, 1) = "0" Then ticket = Mid(ticket, 2)
    If Left(acc, 1) = "0" Then acc = Mid(acc, 2)
        
    ' Initialize Outlook application
    Set olApp = GetOutlookApp()
    If olApp Is Nothing Then Exit Sub

    ' Get the Outlook Namespace
    Set olNamespace = olApp.GetNamespace("MAPI")

    ' Access the shared calendar folder
    Set olRecip = olNamespace.CreateRecipient("NetworkActivationsCalendar")
    If Not olRecip.Resolve Then
        MsgBox "Failed to resolve calendar recipient."
        Exit Sub
    End If

    Set olCalFolder = olNamespace.GetSharedDefaultFolder(olRecip, 9) ' 9 represents olFolderCalendar

    ' Create a search filter for partial match in the subject
   'searchFilter = "@SQL=" & Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%" & pwo & "%'"
   searchFilter = "@SQL=" & _
                    Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%" & ticket & "%' " & _
                    "AND " & _
                    Chr(34) & "urn:schemas:httpmail:subject" & Chr(34) & " like '%" & acc & "%'"


    ' Search for the item
    Set foundItems = olCalFolder.Items.Restrict(searchFilter)

    ' Check if any items were found
    If foundItems.Count > 0 Then
        For Each item In foundItems
            If item.Class = 26 Then ' 26 refers to a meeting item
                item.Display ' Open the meeting
                Exit For
            End If
        Next item
    Else
        MsgBox "No meeting found with the subject containing " & pwo
    End If

    ' Update clipboard with text
    Set clipboardData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    clipboardData.SetText "Received"
    clipboardData.PutInClipboard

    ' Clean up
    Set clipboardData = Nothing
    Set olApt = Nothing
    Set olCalFolder = Nothing
    Set olRecip = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    If GetOutlookApp Is Nothing Then
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function



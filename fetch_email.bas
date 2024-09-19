Sub FetchSpecificEmails()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim mailItem As Object
    Dim items As Object
    Dim i As Integer
    Dim subjectFilter As String
    
    ' Set the subject filter
    subjectFilter = "Coaching"
    
    ' Initialize Outlook application
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation
        Exit Sub
    End If
    
    ' Get MAPI namespace
    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(6)
    
    ' Get all items in the Inbox
    Set items = olFolder.Items
    items.Sort "[ReceivedTime]", False ' Sort by received time descending
    
    ' Loop through items
    For i = 1 To items.Count
        If items(i).Class = 43 Then ' Check if the item is a MailItem
            Set mailItem = items(i)
            If InStr(1, mailItem.Subject, subjectFilter, vbTextCompare) > 0 Then
                ' Display the subject
                MsgBox "Found email: " & mailItem.Subject, vbInformation
                Exit For
            End If
        End If
    Next i
    
    ' Clean up
    Set mailItem = Nothing
    Set items = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

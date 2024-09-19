Sub TestOutlookConnection()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim mailItem As Object
    
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
    
    ' Get the Inbox folder
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 refers to olFolderInbox
    
    ' Get the first email in the Inbox
    Set mailItem = olFolder.Items.GetFirst
    
    If Not mailItem Is Nothing Then
        ' Display the subject of the email
        MsgBox "Subject of the first email: " & mailItem.Subject, vbInformation
    Else
        MsgBox "No emails found in the Inbox.", vbInformation
    End If
    
    ' Clean up
    Set mailItem = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

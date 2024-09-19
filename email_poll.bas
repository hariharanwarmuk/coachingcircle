Sub ExportEmailsFromOutlook()
    Dim olApp As Object
    Dim olNamespace As Object
    Dim olFolder As Object
    Dim olItem As Object
    Dim fs As Object
    Dim jsonFile As Object
    Dim emailData As String
    Dim emailCount As Integer

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

    Set olNamespace = olApp.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(6) ' 6 refers to Inbox

    emailData = "["
    emailCount = 0

    For Each olItem In olFolder.Items
        If olItem.Class = 43 Then ' 43 corresponds to MailItem
            ' Filter emails by subject and sender
            If InStr(1, olItem.Subject, "Coaching", vbTextCompare) > 0 And _
               olItem.SenderEmailAddress = "coach@example.com" Then
                If emailCount > 0 Then
                    emailData = emailData & ","
                End If
                emailData = emailData & "{""subject"":""" & Replace(olItem.Subject, """", "'") & """,""body"":""" & Replace(olItem.Body, """", "'") & """}"
                emailCount = emailCount + 1
            End If
        End If
    Next olItem

    emailData = emailData & "]"

    ' Save the JSON data to emails.json
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set jsonFile = fs.CreateTextFile("D:\work\coaching_circle\emails.json", True)
    jsonFile.WriteLine emailData
    jsonFile.Close

    MsgBox "Emails exported successfully!", vbInformation
End Sub

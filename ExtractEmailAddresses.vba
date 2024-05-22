Sub ExtractEmailAddresses()
    Dim objNamespace As NameSpace
    Dim objFolder As MAPIFolder
    Dim objItem As Object
    Dim objMail As MailItem
    Dim colItems As Items
    Dim i As Integer
    Dim emailAddresses As String
    Dim emailBody As String
    Dim uniqueEmails As Collection
    
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderSentMail)
    Set colItems = objFolder.Items
    Set uniqueEmails = New Collection
    
    On Error Resume Next
    
    For i = colItems.Count To 1 Step -1
        Set objItem = colItems.Item(i)
        If TypeOf objItem Is MailItem Then
            Set objMail = objItem
            ' Extract email addresses from To, CC, BCC fields
            emailAddresses = objMail.To & ";" & objMail.CC & ";" & objMail.BCC
            Call AddUniqueEmails(uniqueEmails, emailAddresses)
            ' Extract email addresses from the email body
            emailBody = objMail.Body
            Call ExtractEmailsFromBody(uniqueEmails, emailBody)
        End If
    Next i
    
    ' Write to a text file
    Dim fs As Object
    Dim aFile As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set aFile = fs.CreateTextFile("C:\Outlook\EmailAddresses.txt", True)
    
    For Each email In uniqueEmails
        aFile.WriteLine email
    Next email
    
    aFile.Close
    
    MsgBox "Email addresses have been extracted to C:\Outlook\EmailAddresses.txt"
End Sub

Sub AddUniqueEmails(col As Collection, emails As String)
    Dim emailArray() As String
    Dim i As Integer
    emailArray = Split(emails, ";")
    
    For i = LBound(emailArray) To UBound(emailArray)
        emailArray(i) = Trim(emailArray(i))
        If IsValidEmail(emailArray(i)) And Not IsInCollection(col, emailArray(i)) Then
            col.Add emailArray(i)
        End If
    Next i
End Sub

Sub ExtractEmailsFromBody(col As Collection, body As String)
    Dim re As Object
    Dim matches As Object
    Dim match As Object
    
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    re.Pattern = "[\w.-]+@[\w.-]+\.[A-Za-z]{2,}"
    
    Set matches = re.Execute(body)
    
    For Each match In matches
        If Not IsInCollection(col, match.Value) Then
            col.Add match.Value
        End If
    Next match
End Sub

Function IsValidEmail(email As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[\w.-]+@[\w.-]+\.[A-Za-z]{2,}$"
    IsValidEmail = re.Test(email)
End Function

Function IsInCollection(col As Collection, email As String) As Boolean
    Dim item As Variant
    On Error Resume Next
    item = col.Item(email)
    If Err.Number = 0 Then
        IsInCollection = True
    Else
        IsInCollection = False
    End If
    On Error GoTo 0
End Function

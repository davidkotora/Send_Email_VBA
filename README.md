# Sent_Email_VBA
Sends an email through code in the VBA editor

Sub SendEmail()
    Dim OutApp As Object
    Dim OutMail As Object
    Dim emailRecipient As String
    Dim emailSubject As String
    Dim emailBody As String
    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Set the email recipient, subject and body
    emailRecipient = "sentemail@example.com"
    emailSubject = "subject@example.com"
    emailBody = "Máš to na stole..."
    
    With OutMail
        .To = emailRecipient
        .Subject = emailSubject
        .Body = emailBody
        .Send
    End With
    
    ' Release the memory of objects
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    ' Display a message after sending the email
    MsgBox "Email sent successfully!"
    
End Sub


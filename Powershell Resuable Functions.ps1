#Powershell Resuable Functions

#set up generic error email to be sent if any errors occur in script
function Send_Error_Email($email_subject, $error_message) {
    Send-MailMessage -From $email_from -To $email_to -Subject $email_subject -Body $error_message -SmtpServer $email_smtp_server -UseSsl
}
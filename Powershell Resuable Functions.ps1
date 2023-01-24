#Powershell Resuable Functions

#set up generic error email to be sent if any errors occur in script
function Send_Error_Email($email_subject, $error_message) {
    Send-MailMessage -From $email_from -To $email_to -Subject $email_subject -Body $error_message -SmtpServer $email_smtp_server -UseSsl
    exit
}

#import csv files, send error email and exit script if there are any errors importing or exporting files
function Import_CSV_File ($file, $file_path) {
    try{$file = Import-Csv - Path $file_path}
    catch{Send_Error_Email}
    return $file
}
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

#check for differences between the reference object/file and the difference object/file
#in this case we are comparing properties of current (active) employees
#Note on SideIndicator
  # <= ... unique to the left-hand side (implied -ReferenceObject parameter)
  # => ... unique to the right-hand side (implied -DifferenceObject parameter)
  # == ... present in both collections (only with -IncludeEqual)
function Check_For_Differences($reference_object, $difference_object) {
    $results = Compare-Object $reference_object $difference_object -Property "First Name", "Last Name", "Employee Number", "Manager Employee Number", "Work Email", "Work Phone", "Mobile Phone", "Job Code", "Department Code", "Hire Date" -Passthru -CaseSensitive | Where-Object{$_.SideIndicator -eq "=>" -and $_.Status -eq "Active"}
    return $results
}

#get array of csv property values
function Get_Array_CSV_Values($file, $property) {
    $csv_values = @()
    foreach($item in $file) {
        $csv_values += $item.$property
    }
    return $csv_values
}

#check for missing or invalid csv values
function Check_For_Invalid_CSV_Values($file, $property, $csv_values) {
    $errors = ""
    foreach($item in $file) {
        if($item.$property -notin $csv_values) {
            $errors += "An error occurred: Invalid or missing $property for $item<br /><br />"
        }
    }
    return $errors
}
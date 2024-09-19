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

#create hashtable so we can look up values by key
function Create_HashTable($file, $key, $value) {
    $hashtable = @{}
    foreach($item in $file) {
        $hashtable[$item.$key] = $item.$value
    }
    return $hashtable
}

#look up the hashtable keys and then send the hashtable values to csv
function Send_HashTable_Values_to_CSV($target_file, $reference_file, $hashtable, $key, $value) {
    foreach($item in $target_file) {
        if($hashtable[$reference_file.$key]) {
            $item.$value = $hashtable[$item.$key]
        }
    }
}

#convert existing dates from format "MM/dd/yyyy" to UTC
#UTC format can be adjusted by changing "yyyy-MM-dd HH:mm:ss.fffffff +00:00"
function Convert_to_UTC($file, $datetime) {
    foreach($item in $file) {
        if($item.$datetime) {
            $item.$datetime = [datetime]::ParseExact($item.$datetime, "MM/dd/yyyy", $null)
            $item.$datetime = $item.$datetime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss.fffffff +00:00")
        }
    }
}

#add additional fields to csv
function Add_Additional_Fields {
    $file | Add-Member -MemberType NoteProperty -Name "FieldName" -Value ""
    $file | Add-Member -MemberType NoteProperty -Name "FieldName1" -Value ""
    $file | Add-Member -MemberType NoteProperty -Name "FieldName2" -Value ""
    $file | Add-Member -MemberType NoteProperty -Name "FieldName3" -Value ""
}

#export to csv button - sample function 
function Export_To_CSV_Button_Click() {
    try {
        $output_file = New-Object psobject
        $output_file | Add-Member -MemberType NoteProperty -Name "WorkEmail" -Value $company_email
        $output_file | Add-Member -MemberType NoteProperty -Name "CompanyID" -Value $company_id
        $output_file | Export-Csv -NoTypeInformation -Path $output_file_path
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("An error occurred: $_" + $_.InvocationInfo.ScriptLineNumber + $_.Exception.StackTrace)
    }
}


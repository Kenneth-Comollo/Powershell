#Powershell Windows Forms

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

#create new csv and button to append items 
$script:output_file_path = "C:\test\NewItemForm_$todays_date.csv"
Set-Content "C:\test\NewItemForm_$todays_date.csv" -Value "WorkEmail, CompanyID"

function Add_Item_To_CSV_Button_Click() {
    $newItem = New-Object psobject
    $newItem | Add-Member -MemberType NoteProperty -Name "WorkEmail" -Value $company_email
    $newItem | Add-Member -MemberType NoteProperty -Name "CompanyID" -Value $company_id
    $newItem | Export-Csv -append -path $output_file_path
    

param (
    [Parameter(Mandatory)] $all_contacts_file, 
    [Parameter(Mandatory)] $extracted_contacts_file, 
    
    $append_to_all_contacts_file = $True,
    $backup_all_contacts_file = $True,

    $export_xlsx = $True, 
    $xlsx_worksheet = "Contacts", 

    $debug_output = $false)

$date = $(get-date -f yyyyMMdd)

$new_contacts_file = "./NewContacts_$date.xlsx"

Write-Host "All Contacts File: $all_contacts_file"
Write-Host "Extracted Contacts File: $extracted_contacts_file"

Write-Host "Backup All Contacts File: $backup_all_contacts_file"
if ($backup_all_contacts_file){
    Copy-Item -Path $all_contacts_file -Destination ($all_contacts_file + "." + $(get-date -f yyyyMMddHHmmss) + ".bak")
}

#$File1 = Import-Csv -Path $all_contacts_file  -Encoding unicode # -Delimiter "`t" 
$File1 = Import-Excel -Path $all_contacts_file -WorksheetName Contacts
 
#$File2 = Import-Csv -Path $extracted_contacts_file  -Encoding unicode # -Delimiter "`t" 
$File2 = Import-Excel -Path $extracted_contacts_file -WorksheetName Contacts

#Compare both files the column Address
$Results = Compare-Object  -ReferenceObject $File1 -DifferenceObject $File2 -Property Address

$Array = @()       
$NewArray = @()       
Foreach($R in $Results)
{
    if( $R.sideindicator -eq "=>" )
    #if( $R.sideindicator -eq "==" )
    {
        $Object = [pscustomobject][ordered] @{
 
            Address = $R.Address
            "Compare indicator" = $R.sideindicator
 
        }
        $Array += $Object
        $NewArray += $File2 | Where-Object {$_.Address -eq $R.Address}
    }
}
 
#Count users in both files
#($Array | sort-object Address | Select-Object * -Unique).count

#Count new users
if ($debug_output) { Write-Host "New Contacts Found : $($Array.Count)"}

#Display results in console
# if ($debug_output) { 
# $Array
# }
# $NewArray | Export-CSV -Path $strFile3 -Encoding unicode # -Append # -Delimiter "`t" 

Write-Host "Append to All Contacts File: $append_to_all_contacts_file"
if ($export_xlsx) {
    if ($append_to_all_contacts_file) {
        $NewArray | Export-Excel -Path $all_contacts_file -AutoSize -TableName $xlsx_worksheet -WorksheetName $xlsx_worksheet -Append
    }
    else {
        Write-Host "New Contacts File: $new_contacts_file"
        $NewArray | Export-Excel -Path $new_contacts_file  -AutoSize -TableName $xlsx_worksheet -WorksheetName $xlsx_worksheet
    }

}

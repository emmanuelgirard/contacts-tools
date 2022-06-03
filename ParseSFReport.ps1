param (
    [Parameter(Mandatory)] $sf_report_file, 
    
    $number_of_won_opportunities = 0,
    $export_xlsx = $True, 
    $xlsx_worksheet = "SF_Accounts_Domains", 

    $debug_output = $false)

$csv_filename = "./SF_Accounts_Domains_$(get-date -f yyyyMMdd).csv"
$xlsx_filename = "./SF_Accounts_Domains_$(get-date -f yyyyMMdd).xlsx"

Write-Host "SF Report File: $sf_report_file"

$allAccounts = Import-Excel -Path $sf_report_file # -WorksheetName Contacts

function splitEmailDomains ($locAccounts) {   
    $arrDomains = @()
    foreach ($anAccount in $locAccounts) {
        $emailDomains = $anAccount."Email Domains"
        if ($emailDomains) {
            $emailDomains = $emailDomains.ToLower().Trim()
        }
        if ($emailDomains -like "*, *") {
            $arrDomains += $emailDomains.split(", ")
        }
        elseif ($emailDomains -like "*,*") {
            $arrDomains += $emailDomains.split(",")
        }
        elseif ($emailDomains -like "*;*") {
            $arrDomains += $emailDomains.Split(";")
        }
        elseif ($emailDomains -like "* *") {
            $arrDomains += $emailDomains.split(" ")
        }
        else {
            $arrDomains += $emailDomains
        }
    }
    return $arrDomains | Sort-Object -Unique
}

#$topAccounts = $allAccounts | Where-Object -Filter {$_."Last Close Date" -ne $null}
$topAccounts = $allAccounts | Where-Object -Filter { $_."Total Won Opportunities (#)" -ge $number_of_won_opportunities }

if ($debug_output) {Write-Host "Email Domain Account: $($topAccounts.Count)"}

$topAccountEmailDomains = @()
$topAccountEmailDomains += splitEmailDomains $topAccounts

$clean_topAccountEmailDomains = @()
foreach($domain in $topAccountEmailDomains){
    Write-Host $domain
    $clean_topAccountEmailDomains += @{Domain = $domain}
}

if (!$export_xlsx) {
    Write-Host "CSV generated file: $csv_filename"
    $clean_topAccountEmailDomains | Select-Object -Property Domain | Export-CSV -Path $csv_filename -Encoding unicode
}

if ($export_xlsx) {
    Write-Host "SF Domains File: $xlsx_filename"
    $clean_topAccountEmailDomains | Select-Object -Property Domain | Export-Excel -Path $xlsx_filename  -AutoSize -TableName $xlsx_worksheet -WorksheetName $xlsx_worksheet
}
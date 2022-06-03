param (
    [Parameter(Mandatory)] $email, 
    $output_file_prefix = "Contacts",
    $folder_patern = "Sent", 
    $max_subfolders_depth = 0,
    $include_from = $false,
    $process_email = $true, 
    $process_meeting = $true, 
    $max_meeting_attendees = 20,
    $from_date = $False, 
    $days = 14, 
    $exclude_domains = @(),
    $exclude_domains_file = "", 
    $include_only_domains_file = "",
    $export_xlsx = $True, 
    $xlsx_worksheet = "Contacts", 
    
    $debug_output = $false)

if (!$from_date) {
    $from_date = ((Get-Date).AddDays(-$days)).ToString("yyyy-MM-dd")
}
$global:api_calls_count = 0

$csv_filename = "./$($output_file_prefix)_$(get-date -f yyyyMMdd).csv"
$xlsx_filename = "./$($output_file_prefix)_$(get-date -f yyyyMMdd).xlsx"

if ($exclude_domains_file -ne "") {
    $exclude_domains_from_file = Import-Csv -Path $exclude_domains_file -Delimiter ","
}
if ($include_only_domains_file -ne "") {
    $include_only_domains_from_file = Import-Csv -Path $include_only_domains_file -Delimiter ","
}

# Install-Module ImportExcel -Scope CurrentUser
# Install-Module Microsoft.Graph -Scope CurrentUser

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Mail
Import-Module Microsoft.Graph.Calendar

Write-Host "Querying Mailbox: $email"
Write-Host "From Date: $from_date"

# This will Pop-Up a Browser to Authenticate to Microsoft and Authorize the PowerShell Application
Connect-MgGraph -Scopes 'User.Read,Mail.Read,Calendars.Read' | Out-Null
if ($debug_output) {$global:api_calls_count += 1}

# Replace with your username (from parameter)
$user_id = $email

Function TrackTime($Time){
    If (!($Time)) { Return Get-Date } Else {
    Return ((get-date) - $Time)
    }
}

# DONE Indent Debug Output
# DONE Number of API Calls in the last exucution
function Get-Childs-Emails ($parent_folder_id, $max, $current = 0) {
    $arrEmails = @()
    $arrEmails += Get-MgUserMailFolderMessage -UserId $user_id -MailFolderId $parent_folder_id -Property "From,ToRecipients,CcRecipients" -Filter "receivedDateTime ge $($from_date)T00:00:00Z" -All
    if ($debug_output) {$global:api_calls_count += 1}

    $subFolders = Get-MgUserMailFolderChildFolder  -UserId $user_id -MailFolderId $parent_folder_id -All
    if ($debug_output) {$global:api_calls_count += 1}
    # if ($debug_output) { 
    #     [console]::CursorLeft = 2*$current
    #     Write-Host "Subfolders: $($subFolders.Count)" 
    # }
    if ($current -lt $max) {
        if ($subFolders.Count -gt 0) {
            # Write-Host "In if subfolders.Count > 0"
            foreach ($subFolder in $subFolders) {
                if ($debug_output) { 
                    [console]::CursorLeft = $current+1
                    # Write-Host "Subfolder: $($subFolder.DisplayName)"
                    Write-Host "$($subFolder.DisplayName)"
                }
                $arrEmails += Get-Childs-Emails $subFolder.Id $max ($current + 1)
                
            }
        }
    }
    return $arrEmails
}

$arrContacts = @()

Write-Host "Process Email: $process_email"
if ($process_email) {
    $time = TrackTime $time
    Write-Host "Max Subfolders : $($max_subfolders_depth)"

    # Get Parent Folder Objet
    $parent_folders = Get-MgUserMailFolder -UserId $user_id -Filter "startswith(DisplayName,'$folder_patern')"
    if ($debug_output) {$global:api_calls_count += 1}
    if ($parent_folders.Count -gt 1) {
        Write-Host "WARNING: More than one folder found with the pattern: $folder_patern" -ForegroundColor Yellow
    }

    $arrEmails = @()
    foreach ($parent_folders_item in $parent_folders) {
        Write-Host "Parent Folder: $($parent_folders_item.DisplayName)"
        $arrEmails += Get-Childs-Emails $parent_folders_item.Id $max_subfolders_depth
    }

    if ($debug_output) { Write-Host "Number of email parsed : $($arrEmails.Count)" }

    Write-Host "Include From: $include_from"

    foreach ($sent_email in $arrEmails) {
        if ($include_from) {
            foreach ($aRecipient in $sent_email.From) {
                try {
                    $aRecipient.EmailAddress.Address = $aRecipient.EmailAddress.Address.ToLower()
                } catch {
                    {
                        Write-Host "WARNING: Recipient with no email address: $($aRecipient.EmailAddress.Name)" -ForegroundColor Yellow
                    }
                }
                $arrContacts += $aRecipient.EmailAddress
            }
        }
        foreach ($aRecipient in $sent_email.ToRecipients) {
            $aRecipient.EmailAddress.Address = $aRecipient.EmailAddress.Address.ToLower()
            $arrContacts += $aRecipient.EmailAddress
        }
        foreach ($aRecipient in $sent_email.CcRecipients) {
            $aRecipient.EmailAddress.Address = $aRecipient.EmailAddress.Address.ToLower()
            $arrContacts += $aRecipient.EmailAddress
        }
    }
    $time = TrackTime $time
    Write-Host "Elapsed Time Parsing Emails: $time"
    $time = $null
}


Write-Host "Process Meeting: $process_meeting"
if ($process_meeting) {
    Write-Host "Max Meeting Attendees: $($max_meeting_attendees)"

    $all_meetings = Get-MgUserEvent -UserId $user_id -Property "Subject,Organizer,Attendees" -Filter "start/dateTime ge '$($from_date)T00:00:00Z'" -All
    if ($debug_output) {$global:api_calls_count += 1}
    # $all_meetings = Get-MgUserEvent -UserId $user_id -Filter "start/dateTime ge '$($from_date)T00:00:00Z'" -All
    # $all_meetings = Get-MgUserEvent -UserId $user_id -Filter "start/dateTime ge '2022-05-01T00:00:00Z'" -All
    if ($debug_output) { Write-Host "Number of meetings : $($all_meetings.Count)" }
}

# Filtering meetings by the number of attendees to isolate internal meetings is the most effective

if ($process_meeting) {
    $time = TrackTime $time
    $arrMeetingSubject = @()
    $meetingParsed = 0
    foreach ($aMeeting in $all_meetings) {
        $am = $false
        $skip = $false
        $arrMeetingSubject += $aMeeting.Subject
        
        if ($aMeeting.Attendees.Count -gt $max_meeting_attendees) {
            if ($debug_output) {Write-Host "WARNING: Meeting with $($aMeeting.Attendees.Count) Attendees: $($aMeeting.Subject)" -ForegroundColor Yellow}
            $skip = $true
        }
        
        foreach ($anAttendee in $aMeeting.Attendees ) {
            if ($anAttendee.EmailAddress.Name.StartsWith("*")) {
                # Write-Host "Skip *"
                $skip = $true
            }
        }
        if (!$skip) {
            # Write-Host "Parsing Meeting Attendees: $($aMeeting.Subject)"

            foreach ($anAttendee in $aMeeting.Attendees ) {
                try {
                    $anAttendee.EmailAddress.Address = $anAttendee.EmailAddress.Address.ToLower()    
                }
                catch {
                        Write-Host "WARNING: Attendee with no email address: $($anAttendee.EmailAddress.Name)" -ForegroundColor Yellow
                }
                $arrContacts += $anAttendee.EmailAddress
            }
            foreach ($anAttendee in $aMeeting.Organizer) {
                try {
                    $anAttendee.EmailAddress.Address = $anAttendee.EmailAddress.Address.ToLower()    
                }
                catch {
                        Write-Host "WARNING: Attendee with no email address: $($anAttendee.EmailAddress.Name)" -ForegroundColor Yellow
                }
                $arrContacts += $anAttendee.EmailAddress
            } 
            $meetingParsed += 1
        }
        else {
            #Write-Host "Skip Meeting: $($aMeeting.Subject)"
                        
        }
    }
    if ($debug_output) { Write-Host "Number of Meetings parsed: $($meetingParsed)"}
    $time = TrackTime $time
    Write-Host "Elapsed Time Parsing Meetings: $time"
    $time = $null
}

$arrContacts = $arrContacts | Sort-Object Address -Unique
$clean_arrContacts = $arrContacts | Select-Object Address, Name, 
@{
    name = 'Domain'
    expr = { $_.Address.Split("@")[1] }
},
@{
    # TODO Gather the different Name Format Extracted from Outlook Display Name
    # This is a First attempt to split the Outlook Display Name into different parts
    # We prefer play safe here and have empty strings instead of messing up our contacts names
    name = 'First'
    expr = {    
       
        # Last, First [9]
        if ($_.Name.Contains('[')) {
            $_.Name.Split(",")[1].Split("[")[0].Trim()
        }
        # First Middle Last (f.last@domain.com)
        elseif ($_.Name.Contains('(') -and $_.Name.Contains('@')) {
            ""
        }
        # Last, First
        elseif ($_.Name.Contains(',')) {
            $_.Name.Split(",")[1].Trim()
        }
        # f.last@domain.com
        elseif ($_.Name.Contains('@')) {
            ""
        }
        # First Last
        else {
            $_.Name.Split(" ")[0].Trim()
        }

    }
},
@{
    name = 'Last'
    expr = {
        if ($_.Name.Contains(',')) {
            $_.Name.Split(",")[0].Trim()
        }
        # First Middle Last (f.last@domain.com)
        elseif ($_.Name.Contains('(') -and $_.Name.Contains('@')) {
            ""
        }        
        elseif ($_.Name.Contains('@')) {
            ""
        }
        else {
            $_.Name.Split(" ")[1].Trim()
        }
       
    }
} | Sort-Object Domain

if ($exclude_domains -gt 0) {
    Write-Host "Excluding Domains from Parameter : "
    $exclude_domains
    $clean_arrContacts = $clean_arrContacts | Where-Object { $_.Domain -notin $exclude_domains }
}

if ($exclude_domains_from_file.Count -gt 0) {
    Write-Host "Excluding Domains from File : "
    foreach ($exclude_domains_item in $exclude_domains_from_file) {
        Write-Host $exclude_domains_item.Domain
    }    
    $clean_arrContacts = $clean_arrContacts | Where-Object { $_.Domain -notin $exclude_domains_from_file.Domain } 
}

# Display domain that were found but excluded because not in the Territory Account Domain List
if ($include_only_domains_from_file.Count -gt 0) {
    $contacts_not_in_the_territory_account_domain_list = $clean_arrContacts | Where-Object { $_.Domain -notin $include_only_domains_from_file.Domain } | Sort-Object Domain -Unique 
    if ($contacts_not_in_the_territory_account_domain_list.Count -gt 0) {
        Write-Host "WARNING: Found Contacts with the following domains but not in the the Include Only Domain File" -ForegroundColor Yellow
        foreach ($contact in $contacts_not_in_the_territory_account_domain_list) {
            Write-Host $contact.Domain
        }
    }
}

if ($include_only_domains_from_file.Count -gt 0) {
    Write-Host "Including Only Contacts from Domains in : $($include_only_domains_file)"
    $clean_arrContacts = $clean_arrContacts | Where-Object { $_.Domain -in $include_only_domains_from_file.Domain }
}

if ($clean_arrContacts.Count -eq 0) {
    Write-Host "No Contacts found" -ForegroundColor Yellow
    exit
}

#Catch File Open 
if (!$export_xlsx) {
    $clean_arrContacts | Export-CSV -Path $csv_filename -Encoding unicode #-Delimiter "`t"
    Write-Host "CSV generated file: $csv_filename"
}

if ($export_xlsx) {
    $clean_arrContacts | Export-Excel -Path $xlsx_filename -AutoSize -TableName $xlsx_worksheet -WorksheetName $xlsx_worksheet
    Write-Host "XLSX generated file: $xlsx_filename"
}

if ($debug_output) { Write-Host "Number of contacts found : $($clean_arrContacts.Count)" }
if ($debug_output) { Write-Host "Number of API calls : $($api_calls_count)" }
 
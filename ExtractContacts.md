# Extract Contacts

This script search for Contacts in O365 using MS Graph API.

It creates an Excel file named *Contacts_YYYYMMDD.xlsx* containing all unique contact found.


## Excel Columns

| Address | Name | Domain | First | Last |
| - | - | - | - | - |
|  |  |  |  |  |

The *Domain* will contain everything after the @ of the contact email. address.
Best effort will be made to fill in the *First* and *Last* column.

## Parameters

| Parameter | Description | Mandatory | Default |
| --------- | ----------- | --------- | ------- |
| $email | Email Adresse of the Account | x | |
| $folder_patern | Folder name | x | "Sent" |
| $max_subfolders_depth | 0 Mean it won't recurse into subfolder, 1 only one level depth ... | x | 0 |
| $include_from | If looking at any other folder then the *Sent Items*, you probably want to enable the From field in the query | x | $False |
| $process_meeting | Parse Calendar Event | | $false |
| $from_date | Search from the date specified in the formal YYYY-MM-DD | | $False |
| $days | Number of days from today to lookup email and meeting for if no $from_date| | 14 |
| $exclude_domains | Email Domain to be ignored | | ("domain1", "domain2") |
| $export_xlsx | Create an Excel File | | $True |
| $xlsx_worksheet | Name of the WorkSheet in Excel | | "Contacts" |
| $debug_output | Display some debug info | | $False |

 
## macOS

```bash
# Install PowerShell

brew install PowerShell
# Start a PowerShell Session
pwsh
```

## Executing the Script

On First Run you will be asked to confirm the installation of the Excel and Graph Modules and also to authenticate using your Microsoft Account to the Graph API.

```powershell
# Extract Contacts found in your Sent Items in the last 14 days
./ExtractContacts.ps1 -email youremailaddress

# Extract Contacts found in your Sent Items in the last 7 days 
# excluding contacts with domains @mydomain and @yourdomain 
./ExtractContacts.ps1 -email youremailaddress -days 7 -exclude_domains ("mydomain","yourdomain")

# Extract Contacts found in your Sent Items since February 1st 2022
./ExtractContacts.ps1 -email youremailaddress -from_date "2022-02-01"
```

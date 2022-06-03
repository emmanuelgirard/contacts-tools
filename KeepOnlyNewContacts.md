# Extract Contacts

This script parse the Excel Extracted Contacts File and search for new Contacts to Append to an Existing Contacts Files.

## Excel Columns

Both files should have at least the following Colums.

| Address | Name | Domain | First | Last |
| - | - | - | - | - |
|  |  |  |  |  |

First line contain the Columns Title.


## Parameters

| Parameter | Description | Mandatory | Default |
| --------- | ----------- | --------- | ------- |
| $all_contacts_file | File that contain all your contacts, it might have more columns and the don't need to be in the same order. | x | |
| $extracted_contacts_file | File that contain new contact found using the Extract Contacts script| x | |
| $append_to_all_contacts_file | Append the New Contacts found in $extracted_contacts_file to $all_contacts_file.  If set to $False, a new file is created *NewContacts_YYYYMMDD.xlsx* | | $True |
| $backup_all_contacts_file | Create a copy of $all_contacts_file before appending New Contacts to it | | $True |
| $export_xlsx | Create an Excel File | | $True |
| $xlsx_worksheet | Name of the WorkSheet in Excel | | "Contacts" |
| $debug_output | Display some debug info | | $false |
 
## macOS

```bash
# Install PowerShell

brew install PowerShell
# Start a PowerShell Session
pwsh
```

```powershell
# Append Only New Contacts found to my Contacts List
./KeepOnlyNewContacts.ps1 -all_contacts_file ./MyListOfContacts.xlsx -extracted_contacts_file ./Contacts_20220530.xlsx

# Create a file containing only the New Contacts found in the last Extracted Contacts Files that arent in my All Contacts File
./KeepOnlyNewContacts.ps1 -all_contacts_file ./MyListOfContacts.xlsx -extracted_contacts_file ./Contacts_20220530.xlsx -append_to_all_contacts_file $False
```
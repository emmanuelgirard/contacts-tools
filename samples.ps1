./ExtractContacts.ps1 -email f.last@domain.com -days 5 -folder_patern "Inbox" -include_from $true -exclude_domains_file ./Exclude_Domains.csv -process_meeting $false -output_file_prefix "_Inbox"
./ExtractContacts.ps1 -email f.last@domain.com -days 5 -folder_patern "Sent" -exclude_domains_file ./Exclude_Domains.csv -process_meeting $false -output_file_prefix "_Sent"
./ExtractContacts.ps1 -email f.last@domain.com -days 5 -folder_patern "Test" -exclude_domains_file ./Exclude_Domains.csv -process_meeting $false -max_subfolders_depth 4 -output_file_prefix "_Test"
./ExtractContacts.ps1 -email f.last@domain.com -days 5 -exclude_domains_file ./Exclude_Domains.csv -process_email $false -process_meeting $true -output_file_prefix "_Meetings"

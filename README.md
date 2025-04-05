# Mails2Excel
This is a PowerAutomate Script to copy content from coherent emails with a certain structure to Excel.

## Instruction
Create a new Flow in Microsoft Power Automate. Copy the code from code.txt to main (just insert from clipboard).

## Adapt
Set the values for the variables
- ***MailBoxName***: Is in outlook the email address (e.g. mail@stuxnerd.org)
- ***MailFolderName***: Is the path in Outlook. It can contain subpaths (e.g. Inbox\MailsForAutomation)
- ***ExcelPath***: Is the path to the Excel file (e.g. C:\DataFromOutlook.xlsx)
  - The file should be empty, but can contain already data (not suggested)
  - Include the ending for the filetype
  - Just tested with absolute paths

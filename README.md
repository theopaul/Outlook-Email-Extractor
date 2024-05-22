# Outlook Email Extractor

This VBA script extracts all email addresses from the "To", "CC", "BCC" fields, and the body of sent emails in Microsoft Outlook. The extracted email addresses are saved to a text file located at `C:\Outlook\EmailAddresses.txt`.

## Usage

1. Open Outlook.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by going to `Insert > Module`.
4. Copy and paste the script from this repository into the module.
5. Press `Alt + F8`, select `ExtractEmailAddresses`, and click "Run".

The email addresses will be extracted and saved to `C:\Outlook\EmailAddresses.txt`.

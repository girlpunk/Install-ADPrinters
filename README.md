# Install-ADPrinters
Flexable way of deploying printers on Windows using PowerShell

Printers are selected by entering Active Directory OUs into the printer's comment field. Computers in the specified OU will have the printer installed.
The printer can be made default by putting an asterisk after the OU name.

The script should be set as a startup script for the PC using Group Policy.
# UpdateWorkDays
Update-WorkDays.ps1 is a a PowerShell script to update Exchange Server mailbox calendar's "WorkDays" property.

# Background
When you use Exchange Server's cmdlet to configure WorkDays to either "Weekdays" or "AllDays", Outlook does not understand these literal values.

e.g.
```PowerShell
Set-MailboxCalendarConfiguration user01 -WorkDays AllDays
```

As a result, the calendar looks "grayed-out".  
This script fixes this situation by modifying the WorkDays value as follows:

- if "AllDays"  --> "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"
- if "Weekdays" --> "Monday Tuesday Wednesday Thursday Friday"

# Requirement
- PowerShell v2 or later
- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

# How to use

1. With Exchange Management Shell, grant ApplicationImpersonation role to a user who runs the script  
   This script uses EWS Impersonation to access the target mailbox.

   ```PowerShell
   New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User contoso\administrator
   ```
   
2. Download EWS Managed API
   
   [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

3. Start PowerShell and dot source the file

   e.g. 
   ```
   . c:\tmp\UpdateWorkDays.ps1
   ```
  
5. Execute Update-WorkDays

   Here are the mandatory parameters:
   
   |name|meagning
   |----|-
   |EwsManagedApiPath|Path to Microsoft.Exchange.WebServices.dll
   |Server|Name of the Exchange server to send EWS request (For Exchange Online, "outlook.office365.com" or "outlook.office.com")
   |Credential|Credential of a user who accesses the target mailbox (a user with ApplicationImpersonation role)
   |TargetMailboxSmtpAddress|Target mailbox whose WorkHous you are trying to fix
   
   Optional parameters:

   |name|meagning
   |----|-
   |EnableTrace|Switch parameter to enable tracing
   |TraceFile|Path to a trace file
   
   
   e.g.
   ```PowerShell
   Update-WorkDays -EwsManagedApiPath "C:\Microsoft.Exchange.WebServices.dll" -Server myExchange.contoso.local -Credential (Get-Credential) -TargetMailboxSmtpAddress user01@contoso.local
   ```
   
   
   
  

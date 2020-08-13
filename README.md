UpdateWorkDays.ps1 is a PowerShell script to update Exchange Server mailbox calendar's "WorkDays" property.
It defines a function `Update-WorkDays`.

[Download](https://github.com/jpmessaging/UpdateWorkDays/releases/download/v1.0/UpdateWorkDays.ps1)

## Background
When you use Exchange Server's cmdlet to configure WorkDays to either "Weekdays" or "AllDays", Outlook does not understand these literal values.

e.g.
```PowerShell
Set-MailboxCalendarConfiguration user01 -WorkDays AllDays
```

As a result, the calendar looks "grayed-out".  
This script fixes this situation by modifying the WorkDays value as follows:

- if "AllDays"  --> "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"
- if "Weekdays" --> "Monday Tuesday Wednesday Thursday Friday"

## Requirement
- PowerShell v2 or later
- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

## How to use

1. With Exchange Management Shell, grant ApplicationImpersonation role to a user who runs the script  
   This script uses EWS Impersonation to access the target mailbox.

   ```PowerShell
   New-ManagementRoleAssignment -Role "ApplicationImpersonation" -User contoso\administrator
   ```
   
2. Download EWS Managed API
   
   [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)

3. Start PowerShell and Import-Module the file

   e.g. 
   ```
   Import-Module C:\tmp\UpdateWorkDays.ps1
   ```
  
5. Execute Update-WorkDays

   Here are the mandatory parameters:
   
   | name                     | meaning                                                                                                                |
   | ------------------------ | ---------------------------------------------------------------------------------------------------------------------- |
   | EwsManagedApiPath        | Path to Microsoft.Exchange.WebServices.dll                                                                             |
   | Server                   | Name of the Exchange server to send EWS request (For Exchange Online, "outlook.office365.com" or "outlook.office.com") |
   | Credential               | Credential of a user who accesses the target mailbox (a user with ApplicationImpersonation role)                       |
   | TargetMailboxSmtpAddress | Target mailbox whose WorkDays you are trying to fix                                                                    |
   
   Optional parameters:

   | name        | meaning                            |
   | ----------- | ---------------------------------- |
   | EnableTrace | Switch parameter to enable tracing |
   | TraceFile   | Path to a trace file               |
   
   
   e.g.
   ```PowerShell
   Update-WorkDays -EwsManagedApiPath "C:\Microsoft.Exchange.WebServices.dll" -Server myExchange.contoso.local -Credential (Get-Credential) -TargetMailboxSmtpAddress user01@contoso.local
   ```
   
## License
Copyright (c) 2020 Ryusuke Fujita

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.   
   
  

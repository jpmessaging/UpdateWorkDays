UpdateWorkDays.psm1 is a PowerShell script to update Exchange Server mailbox calendar's "WorkDays" property.  
It defines functions `Update-WorkDays` and `Get-Token`.  

[Download](https://github.com/jpmessaging/UpdateWorkDays/releases/download/v2020-11-03/UpdateWorkDays.zip)

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
- PowerShell v3 or later  
- .NET Framework 4.6.1 or later  

The following modules need to be placed under "modules" sub folder where the script file is. These are included in the release package.  

- [Microsoft Exchange Web Services Managed API 2.2](https://www.microsoft.com/en-us/download/details.aspx?id=42951)  
- [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)  
- [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)  

## How to use  

1. With Exchange Management Shell, grant ApplicationImpersonation role to a user who runs the script  

   This script uses EWS Impersonation to access the target mailbox.  

   e.g.  
   ```PowerShell  
   New-ManagementRoleAssignment -Role 'ApplicationImpersonation' -User contoso\administrator  
   ```  

2. Start PowerShell and Import-Module the script  

   e.g.  
   ```PowerShell  
   Import-Module 'C:\tmp\UpdateWorkDays.psm1'  
   ```  

3. If you are using Modern Auth, get a token  

   You can use `Get-Token` included in this script.  

   e.g.  
   ```PowerShell  
   $token = Get-Token -ClientId '63ce5cc6-c944-4baa-83d1-5cac8cdf487e' -Scopes 'https://outlook.office365.com/EWS.AccessAsUser.All'  
   ```  

   \* Client ID above is just an example. Use your own application's Client ID.  

4. Execute Update-WorkDays  

   #### mandatory parameters:  

   | name                     | meaning                                                                   |
   | ------------------------ | ------------------------------------------------------------------------- |
   | Server                   | Name of the Exchange server. For Exchange Online, `outlook.office365.com` |
   | TargetMailboxSmtpAddress | Target mailbox whose WorkDays you are trying to fix                       |

   #### Conditionally mandatory parameters:  
   \* These parameters are mutually exclusive: `Credential` for legacy auth and `Token` for Modern auth.  

   | name       | meaning                           |
   | ---------- | --------------------------------- |
   | Credential | Credential used for Legacy auth   |
   | Token      | Access token used for Modern auth |

   ####  Optional parameters:  

   | name              | meaning                                    |
   | ----------------- | ------------------------------------------ |
   | EwsManagedApiPath | Path to Microsoft.Exchange.WebServices.dll |
   | EnableTrace       | Switch parameter to enable tracing         |
   | TraceFile         | Path to a trace file                       |


   e.g.  
   ```PowerShell  
   Update-WorkDays -Server 'myExchange.contoso.local' -TargetMailboxSmtpAddress 'user01@contoso.local' -Credential (Get-Credential) -EnableTrace -TraceFile 'C:\temp\trace.txt'  
   ```  

   e.g.  
   \* The following Client ID is just an example. Use your own application's Client ID.  
   \* If your application is registered as a muti-tenant app, then you do not need to provide `TenantId` parameter for `Get-Token` (Just as this example). If your application is a single-tenant app, then specify the tenant name or GUID in `TenantId`.  

   ```PowerShell  
   $token = Get-Token -ClientId '63ce5cc6-c944-4baa-83d1-5cac8cdf487e' -Scopes 'https://outlook.office365.com/EWS.AccessAsUser.All'  
   Update-WorkDays -Server 'outlook.office365.com' -TargetMailboxSmtpAddress 'user01@contoso.com' -Token $token.AccessToken -EnableTrace -TraceFile 'C:\temp\trace.txt'  
   ```  

## About Modern Auth
If you are using Modern Auth, then you first need to register an application in your Azure AD.


- `Get-Token` uses `https://login.microsoftonline.com/common/oauth2/nativeclient` as the default Redirect URI. If your application uses a different URI, provide it to `RedirectUri`.
- For API Permissions, add `Exchange`'s `EWS.AccessAsUser.All`

### Reference
- [Authenticate an EWS application by using OAuth](https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

## License  
Copyright (c) 2020 Ryusuke Fujita  

This software is released under the MIT License.  
http://opensource.org/licenses/mit-license.php  

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:  

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.  

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.  
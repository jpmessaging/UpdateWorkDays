#Requires -Version 3

$TraceListenerDefinition = @"
using System;
using System.Text;
using System.IO;
using Microsoft.Exchange.WebServices.Data;

public class TraceListener : ITraceListener, IDisposable
{
        private StreamWriter writer;
        private int bufferSize = 80*1024; // 80 KB

        public TraceListener(string filePath)
        {
            writer = new StreamWriter(filePath, true, Encoding.UTF8, bufferSize);
        }

        // this doesn't flush every write. So make sure to Dispose.
        public void Trace(string traceType, string traceMessage)
        {
            writer.Write(traceMessage);
        }

        private bool disposed = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                writer.Dispose();
                GC.SuppressFinalize(this);
            }

            disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
        }

}
"@

$TrustAllCertificatePolicyDefinition = @"
using System.Net;
using System.Security.Cryptography.X509Certificates;

public class TrustAllCertsPolicy : ICertificatePolicy
{
    public bool CheckValidationResult(ServicePoint srvPoint, X509Certificate certificate, WebRequest request, int certificateProblem)
    {
        return true;
    }
}
"@


<#
.SYNOPSIS
This function updates "WorkDays" of mailbox's calendar configuration according to the format Outlook understands.

- If it's "Weekdays", it'll be updated to "Monday Tuesday Wednesday Thursday Friday"
- If it's "AllDays", it'll be updated to "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"

The calendar configuration is stored in IPM.Configuration.WorkHours item's PR_ROAMING_XMLSTREAM.

.DESCRIPTION

.PARAMETER Server
Mandatory name of the EWS server (e.g. outlook.office.com).  Note that you don't need the full URL of EWS.

.PARAMETER TargetMailboxSmtpAddress
Mandatory SMTP address of the target mailbox.  This is the mailbox to update calendar configuration.

.PARAMETER Credential
PSCredential. This is used to access the target mailbox with legacy auth.
This parameter is mutually exclusive with Token parameter.

If you use impersonation to access a mailbox, make sure to give the accout ApplicationImpersonation RBAC role.

e.g. Giving ApplicationImpersonation role to administrator
New-ManagementRoleAssignment -Role ApplicationImpersonation -User administrator

.PARAMETER Token
Token for Modern Authentication.
This parameter is mutually exclusive with Credential parameter.

You can use Get-Token in this script to obtain a token.

   e.g.
   $token = Get-Token -ClientId <Client ID of application> -Scopes "https://outlook.office365.com/EWS.AccessAsUser.All"

.PARAMETER EwsManagedApiPath
Optional path for EWS Managed API DLL (i.e. Microsoft.Exchange.WebServices.dll)
By default, it looks under "modules" folder in the script's directory.

.PARAMETER EnableTrace
Optional switch to enable a trace.

.PARAMETER TraceFile
Optional path for the trace. If EnableTrace is true but TraceFile is not specified, then the trace will be output on the console.

.PARAMETER TrustAllCertificate
Optional switch to ignore any certificate issues.

.EXAMPLE
Use Modern auth to update room01@contoso.com's WorkDays in Exchange Online.

    $token = Get-Token -ClientId 63ce5cc6-c944-4baa-83d1-5cac8cdf487e -Scopes "https://outlook.office365.com/EWS.AccessAsUser.All"
    Update-WorkDays -Server outlook.office365.com -TargetMailboxSmtpAddress room01@contoso.com -Token $token.AccessToken -EnableTrace -TraceFile C:\temp\trace.txt

.EXAMPLE
Use Legacy auth to update room01@contoso.com's WorkDays in Exchange Online.

    Update-WorkDays -Server outlook.office365.com -TargetMailboxSmtpAddress room01@contoso.com -Credential (Get-Credential contoso\administrator) -EnableTrace -TraceFile C:\temp\trace.txt

.EXAMPLE
Use Legacy auth to update room01@contoso.com's WorkDays in Exchange Online. Also use the EWS DLL specified by EwsManagedApiPath.

    Update-WorkDays -EwsManagedApiPath "C:\temp\Microsoft.Exchange.WebServices.dll" -Server outlook.office365.com -TargetMailboxSmtpAddress room01@contoso.com -Credential (Get-Credential contoso\administrator) -EnableTrace -TraceFile C:\temp\trace.txt

.EXAMPLE
Update room01@contoso.com's WorkDays in OnPremise Exchange.

    Update-WorkDays -Server myExchange.contoso.com -TargetMailboxSmtpAddress room01@contoso.com -Credential (Get-Credential contoso\administrator) -EnableTrace -TraceFile C:\temp\trace.txt

.OUTPUTS
It outputs a PSCustomObject with the following properties:

- Mailbox : Target mailbox. Same as TargetMailboxSmtpAddress input parameter
- Server : EWS server. Same as Server input parameter
- Updated : Boolean indicating if the calendar configuration was updated or not.
- WorkHoursBeforeUpdate : WorkHours configuration before
- WorkHoursAfterUpdate : WorkHours configuration after

.NOTES
See https://github.com/jpmessaging/UpdateWorkDays

Copyright 2020 Ryusuke Fujita

This software is released under the MIT License.
http://opensource.org/licenses/mit-license.php

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

function Update-WorkDays
{
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)]
    $Server,
    [Parameter(Mandatory=$true)]
    $TargetMailboxSmtpAddress,
    [Parameter(ParameterSetName='LegacyAuth')]
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter(ParameterSetName='ModernAuth')]
    [string]$Token,
    [string]$EwsManagedApiPath,
    [switch]$EnableTrace,
    $TraceFile,
    [switch]$TrustAllCertificate
    )

    Write-Verbose "Loading EWS Managed API"
    try {
        if (-not $EwsManagedApiPath) {
            $EwsManagedApiPath = Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Exchange.WebServices.dll'
        }

        Add-Type -Path $EwsManagedApiPath
    }
    catch {
        Write-Error "Failed to load EWS Managed API`n$_"
        return
    }

    $ews = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $ews.Url = [URI]"https://$Server/ews/Exchange.asmx"
    $ews.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailboxSmtpAddress)
    #$ews.PreferredCulture = New-Object System.Globalization.CultureInfo -ArgumentList "ja-JP"

    switch -Wildcard ($PSCmdlet.ParameterSetName) {
        'LegacyAuth' {
            Write-Verbose "Credential is provided. Use it for legacy auth"
            $ews.Credentials = $Credential.GetNetworkCredential()
            break
        }

        'ModernAuth' {
            Write-Verbose "Token is provided. Use it for modern auth"
            $ews.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($Token)
            break
        }
    }

    if ($EnableTrace) {
        Write-Verbose "Enabling trace"
        $ews.TraceEnabled = $true
        $ews.TraceEnablePrettyPrinting = $true
        $ews.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All

        if ($TraceFile) {
            if (-not ("TraceListener" -as [type])) {
                Add-Type -ReferencedAssemblies $EwsManagedApiPath -TypeDefinition $TraceListenerDefinition -Language CSharp
            }
            $ews.TraceListener = New-Object TraceListener -ArgumentList $TraceFile
        }
    }

    if ($TrustAllCertificate) {
        if (-not ("TrustAllCertsPolicy" -as [type])) {
            Add-Type $TrustAllCertificatePolicyDefinition
        }

        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    }

    try {
        Write-Verbose "Sending a GetFolder request for Calendar folder ..."
        $calFolder = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($ews, ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar), [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)

        Set-Variable WorkHoursItemClass -Option ReadOnly 'IPM.Configuration.WorkHours'
        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList 1
        $view.PropertySet = [Microsoft.Exchange.WebServices.Data.PropertySet]::IdOnly
        $view.Traversal = [Microsoft.Exchange.WebServices.Data.ItemTraversal]::Associated
        $searchFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo -ArgumentList ([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass), $WorkHoursItemClass

        Write-Verbose "Sending a FineItem request ..."
        $findResults = $calFolder.FindItems($searchFilter, $view);

        # there should be at most one item with "IPM.Configuration.WorkHours" class
        if ($findResults.TotalCount -ne 1)
        {
            Write-Warning "FindItem found $($findResults.TotalCount) item with $WorkHoursItemClass."
            return;
        }

        $item = $findResults | Select-Object -First 1

        Write-Verbose "Sending a GetItem request to grab PR_ROAMING_XMLSTREAM ..."
        $PR_ROAMING_XMLSTREAM = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition -ArgumentList 0x7C08, ([Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly), $PR_ROAMING_XMLSTREAM
        $workHoursItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($ews, $item.Id, $propSet);
        $roamingXmlProperty = $workHoursItem.ExtendedProperties | Select-Object -First 1

        if (-not $roamingXmlProperty)
        {
            Write-Warning "The item doesn't have PR_ROAMING_XMLSTREAM"
            return
        }

        $bytes = $roamingXmlProperty.Value -as [byte[]]
        if (-not $bytes)
        {
            Write-Warning "PR_ROAMING_XMLSTREAM doesn't have value"
            return
        }

        Write-Verbose "Parsing xml and will update WorkDays value if necessary"
        [xml]$xmldoc = [Text.Encoding]::ASCII.GetString($bytes)
        $workDaysElement = $xmldoc.GetElementsByTagName('WorkDays') | Select-Object -First 1

        $result = New-Object PSCustomObject -Property @{
            Mailbox = $TargetMailboxSmtpAddress
            Server = $Server
            Updated = $false
            WorkHoursBeforeUpdate = $xmldoc.OuterXml
            WorkHoursAfterUpdate = $null
        }

        $needUpdate = $false
        if ($workDaysElement -and $workDaysElement.InnerText -eq 'Weekdays')
        {
            # update to "Monday Tuesday Wednesday Thursday Friday"
            $workDaysElement.InnerText = 'Monday Tuesday Wednesday Thursday Friday'
            $needUpdate = $true;
        }
        elseif ($workDaysElement -and $workDaysElement.InnerText -eq 'AllDays')
        {
            # update to "Sunday Monday Tuesday Wednesday Thursday Friday Saturday"
            $workDaysElement.InnerText = 'Sunday Monday Tuesday Wednesday Thursday Friday Saturday'
            $needUpdate = $true;
        }

        if ($needUpdate)
        {
            $workHoursItem.SetExtendedProperty($PR_ROAMING_XMLSTREAM, ([Text.Encoding]::ASCII.GetBytes($xmldoc.OuterXml)))
            Write-Verbose "Sending an UpdateItem request ..."
            $workHoursItem.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite)
            $result.WorkHoursAfterUpdate = $xmldoc.OuterXml
            $result.Updated = $true
        }

        Write-Output $result
    }
    catch
    {
        Write-Error "Failed to update the property`n$_"
    }
    finally
    {
        if ($ews.TraceListener -as [IDisposable])
        {
            $ews.TraceListener.Dispose()
        }
    }
}

<#
.SYNOPSIS
This function returns an instance of Microsoft.Identity.Client.LogCallback delegate which calls the given scriptblock when LogCallback is invoked.
#>
function New-LogCallback {
    [CmdletBinding()]
    param (
    # Scriptblock to be called when MSAL invokes LogCallback
    [Parameter(Mandatory=$true)]
    [scriptblock]$Callback,

    # Remaining arguments to be passd to Callback scriptblock via $Event.MessageData
    [Parameter(ValueFromRemainingArguments)]
    [object[]]$ArgumentList
    )

    # Class that exposes an event of type Microsoft.Identity.Client.LogCallback that Register-ObjectEvent can register to.
    $LogCallbackProxyType = @"
        using System;
        using System.Threading;
        using Microsoft.Identity.Client;

        public sealed class LogCallbackProxy
        {
            // This is the exposed event. The sole purpose is for Register-ObjectEvent to hook to.
            public event LogCallback Logging;

            // This is the LogCallback delegate instance.
            public LogCallback Callback
            {
                get { return new LogCallback(OnLogging); }
            }

            // Raise the event
            private void OnLogging(LogLevel level, string message, bool containsPii)
            {
                LogCallback temp = Volatile.Read(ref Logging);
                if (temp != null) {
                    temp(level, message, containsPii);
                }
            }
        }
"@

    if (-not ("LogCallbackProxy" -as [type])) {
        Add-Type $LogCallbackProxyType -ReferencedAssemblies (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
    }

    $proxy = New-Object LogCallbackProxy
    Register-ObjectEvent -InputObject $proxy -EventName Logging -Action $Callback -MessageData $ArgumentList | Out-Null

    $proxy.Callback
}

<#
.SYNOPSIS
Obtains a modern auth token (maybe from a cached one if available).

.NOTES
You need the following MSAL.NET modules under "modules" sub folder:

 [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
 [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)

 Folder structure should look like this:

    SomeFolder
    |  UpdateWorkDays.psm1
    |
    |- modules
          Microsoft.Identity.Client.dll
          Microsoft.Identity.Client.Extensions.Msal.dll

.LINK
[AzureAD/microsoft-authentication-library-for-dotnet](https://github.com/AzureAD/microsoft-authentication-library-for-dotnet)
[AzureAD/microsoft-authentication-extensions-for-dotnet](https://github.com/AzureAD/microsoft-authentication-extensions-for-dotnet)

#>
function Get-Token {
    [CmdletBinding()]
    param(
    # Client ID (Application ID) of the registered application.
    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    # Tenant ID. By default, it uses '/common' endpoint for multi-tenant app. For a single-tenant app, specify the tenant name or GUID (e.g. "contoso.com", "contoso.onmicrosoft.com", "333b3ed5-0ac4-4e75-a1cd-db9e8f593ff3")
    [string]$TenantId = 'common',

    # Array of scopes to request.  By default, "openid", "profile", and "offline_access" are included.
    [string[]]$Scopes,

    # Refirect URI for the application. When this is not given, "https://login.microsoftonline.com/common/oauth2/nativeclient" will be used.
    # Make sure to use the same URI as the one registered for the application.
    [string]$RedirectUri,

    # Clear the cached token and force to get a new token.
    [switch]$ClearCache,

    # Enable MSAL logging. Log file will be msal.log under the script folder.
    [switch]$EnableLogging
    )

    # Need MSAL.NET DLL under modules
    # https://github.com/AzureAD/microsoft-authentication-library-for-dotnet
    # [MSAL.NET](https://www.nuget.org/packages/Microsoft.Identity.Client)
    if (-not ('Microsoft.Identity.Client.AuthenticationResult' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.dll')
        }
        catch {
            Write-Error $_
            return
        }
    }

    # [MSAL.NET Extensions](https://www.nuget.org/packages/Microsoft.Identity.Client.Extensions.Msal/)
    if (-not ('Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper' -as [type])) {
        try {
            Add-Type -Path (Join-Path (Split-Path $PSCommandPath) 'modules\Microsoft.Identity.Client.Extensions.Msal.dll')
        }
        catch {
            Write-Error $_
            return
        }
    }

    # Configure & create a PublicClientApplication
    $builder = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($ClientId).WithAuthority("https://login.microsoftonline.com/$TenantId/")

    if ($RedirectUri) {
        $builder.WithRedirectUri($RedirectUri) | Out-Null
    }
    else {
        # WithDefaultRedirectUri() makes the redirect_uri "https://login.microsoftonline.com/common/oauth2/nativeclient".
        # Without it, redirect_uri would be "urn:ietf:wg:oauth:2.0:oob".
        $builder.WithDefaultRedirectUri() | Out-Null
    }

    $writer = $null

    if ($EnableLogging) {
        $logFile = Join-Path (Split-Path $PSCommandPath) 'msal.log'
        [IO.StreamWriter]$writer = [IO.File]::AppendText($logFile)
        Write-Verbose "MSAL Logging is enabled. Log file: $logFile"

        # Add a CSV header line
        $writer.WriteLine("datetime,level,containsPii,message");

        $builder.WithLogging(
            # Microsoft.Identity.Client.LogCallback
            (New-LogCallback {
                param([Microsoft.Identity.Client.LogLevel]$level, [string]$message, [bool]$containsPii)

                $writer = $Event.MessageData[0]
                $writer.WriteLine("$((Get-Date).ToString('o')),$level,$containsPii,`"$message`"")

            } -ArgumentList $writer),

            [Microsoft.Identity.Client.LogLevel]::Verbose,
            # enablePiiLogging
            $true,
            # enableDefaultPlatformLogging
            $false
        ) | Out-Null
    }

    $publicClient = $builder.Build()

    # Configure caching
    $cacheFileName = "msalcache.bin"
    $cacheDir = Split-Path $PSCommandPath
    $storagePropertiesBuilder = New-Object Microsoft.Identity.Client.Extensions.Msal.StorageCreationPropertiesBuilder($cacheFileName, $cacheDir, $ClientId)
    $storageProperties = $storagePropertiesBuilder.Build()
    $cacheHelper = [Microsoft.Identity.Client.Extensions.Msal.MsalCacheHelper]::CreateAsync($storageProperties).GetAwaiter().GetResult()
    $cacheHelper.RegisterCache($publicClient.UserTokenCache)

    if ($ClearCache) {
        $cacheHelper.Clear()
    }

    # Get an account
    $firstAccount = $publicClient.GetAccountsAsync().GetAwaiter().GetResult() | Select-Object -First 1

    # By default, MSAL asks for scopes: openid, profile, and offline_access.
    try {
        $publicClient.AcquireTokenSilent($Scopes, $firstAccount).ExecuteAsync().GetAwaiter().GetResult()
    }
    catch [Microsoft.Identity.Client.MsalUiRequiredException] {
        try {
            $publicClient.AcquireTokenInteractive($Scopes).ExecuteAsync().GetAwaiter().GetResult()
        }
        catch {
            Write-Error $_
        }
    }
    catch {
        Write-Error $_
    }
    finally {
        if ($writer){
            $writer.Dispose()
        }
    }
}


Export-ModuleMember -Function Update-WorkDays, Get-Token
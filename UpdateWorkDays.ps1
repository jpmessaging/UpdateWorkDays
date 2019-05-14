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

        // this doesn't flush everyt write. So make sure to Dispose.                     
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

.PARAMETER EwsManagedApiPath
Mandatory path for EWS Managed API dll (i.e. Microsoft.Exchange.WebServices.dll)

.PARAMETER Server
Mandatory name of the EWS server (e.g. outlook.office.com).  Note that you don't need the full URL of EWS.

.PARAMETER Credential
Mandatory PSCredential. This is used to access the target mailbox.
If you use impersonation to access a mailbox, make sure to give the accout ApplicationImpersonation RBAC role.

e.g. Giving ApplicationImpersonation role to administrator
New-ManagementRoleAssignment -Role ApplicationImpersonation -User administrator

.PARAMETER TargetMailboxSmtpAddress
Mandatory SMTP address of the target mailbox.  This is the mailbox to update calendar configuration.

.PARAMETER EnableTrace
Optional switch to enable a trace.

.PARAMETER TraceFile
Optional path for the trace. If EnableTrace is true but TraceFile is not specified, then the trace will be output on the console.

.PARAMETER TrustAllCertificate
Optional switch to ignore any certificate issues.

.EXAMPLE
Update-WorkDays -EwsManagedApiPath "C:\temp\Microsoft.Exchange.WebServices.dll" -Server outlook.office.com -Credential (Get-Credential contoso\administrator) -TargetMailboxSmtpAddress room01@contoso.local -EnableTrace -TraceFile C:\temp\trace.txt

.OUTPUTS
It outputs a PSCustomObject with the following properties:

- Mailbox : Target mailbox. Same as TargetMailboxSmtpAddress input parameter
- Server : EWS server. Same as Server input parameter
- Updated : Boolean indicating if the calendar configuration was updated or not.
- WorkHoursBeforeUpdate : WorkHours configuration before
- WorkHoursAfterUpdate : WorkHours configuration after

.NOTES
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>
function Update-WorkDays
{   
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)]
    $EwsManagedApiPath,
    [Parameter(Mandatory=$true)]
    $Server,    
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential] $Credential,
    [Parameter(Mandatory=$true)]
    $TargetMailboxSmtpAddress,
    [switch]$EnableTrace,
    $TraceFile,
    [switch]$TrustAllCertificate
    )       

    try
    {
        Add-Type -Path $EwsManagedApiPath
        
    }
    catch
    {
        Write-Error "Failed to load EWS Managed API`n$_"        
        return
    }

    $ews = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)
    $ews.Url = [URI]"https://$Server/ews/Exchange.asmx"
    $ews.Credentials = $Credential.GetNetworkCredential()
    $ews.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailboxSmtpAddress)
    #$ews.PreferredCulture = New-Object System.Globalization.CultureInfo -ArgumentList "ja-JP"

    if ($EnableTrace)
    {        
        $ews.TraceEnabled = $true
        $ews.TraceEnablePrettyPrinting = $true
        $ews.TraceFlags = [Microsoft.Exchange.WebServices.Data.TraceFlags]::All        

        if ($TraceFile)
        {             
            if (-not("TraceListener" -as [type]))
            {
                Add-Type -ReferencedAssemblies $EwsManagedApiPath -TypeDefinition $TraceListenerDefinition -Language CSharp
            }
            $ews.TraceListener = New-Object TraceListener -ArgumentList $TraceFile            
        }
    }

    if ($TrustAllCertificate)
    {
        if (-not ("TrustAllCertsPolicy" -as [type]))
        {
            Add-Type $TrustAllCertificatePolicyDefinition            
        }

        [System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy
    }

    try
    {
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

        $item = $findResults | select -First 1

        Write-Verbose "Sending a GetItem request to grab PR_ROAMING_XMLSTREAM ..."
        $PR_ROAMING_XMLSTREAM = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition -ArgumentList 0x7C08, ([Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
        $propSet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet -ArgumentList ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly), $PR_ROAMING_XMLSTREAM
        $workHoursItem = [Microsoft.Exchange.WebServices.Data.Item]::Bind($ews, $item.Id, $propSet);        
        $roamingXmlProperty = $workHoursItem.ExtendedProperties | select -First 1

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
        $workDaysElement = $xmldoc.GetElementsByTagName('WorkDays') | select -First 1

        $result = New-Object PSCustomObject -Property @{Mailbox = $TargetMailboxSmtpAddress
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
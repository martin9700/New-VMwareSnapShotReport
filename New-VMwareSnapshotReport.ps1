<#
.SYNOPSIS
    Simple HTML report showing all snapshots on your VMware vSphere
    environment
.DESCRIPTION
    Run this script to produce a simple HTML report showing all the 
    snapshots in your vSphere environment.  This script is designed to
    run without prompting you for credentials, saving the password in
    a (mostly) secure, encrypted file.  You must run the script once 
    manually to produce the proper credential file before putting it 
    into a scheduled task.
    
    Requires VMware PowerCLI be installed.

    How to use:
    1. Log into the server where you intend to run this script as a Scheduled Task. You must log in using the service
       account name and password you will be using in the Task Scheduler.
    2. Modify the default parameter values to match your environment.  Do not put in anything for $Credential.  Fields
       you should modify are:  VICenter, To, From and SMTPServer.  You can change Path if you want to save your reports
       in an alternative location, default is a Reports subfolder where you have the script stored.
    3. Run the script.  It will prompt you for the username and password to authenticate to your VCenter.
    4. Create your scheduled task.
    
.PARAMETER VIServer
    Name of your vSphere vCenter server, or the name of your ESXi host.

.PARAMETER Credential
    PSCredential to authenticate to VIServer.  If you don't specify the script will prompt you for the credential and then
    save it in vcenter.xml.  To run this properly you must be logged in using the service account name of that will be running
    this script in the Task Scheduler.

.PARAMETER Path
    Path where the HTML report will be saved.

.PARAMETER To
    Who the emailed report is going to

.PARAMETER From
    Who the emailed report is coming from

.PARAMETER SMTPServer
    The IP address or name of the SMTP relay you want to use

.EXAMPLE
    .\New-VMwareSnapshotReport.ps1 -VIServer VCenter1 -Path \\server\share\reports
    
    Create a report of all the snapshots of VM's under the control of the
    VCenter1 vCenter server.  The resulting HTML report will be
    saved on the same server and share, but in the directory "reports".

.EXAMPLE
    .\New-VMwareSnapshotReport.ps1 -VIServer VCenter1 -Path \\server\share\reports -To "me@mydomain.com" -From "you@yourdomain.com" -SMTPServer "MyExchange1"
    
    Same as the example above, but overriding the default mailing parameters
    to send to me@mydomain.com, from you@yourdomain.com and using the MyExchange1
    server to relay the email.

.INPUTS
    None
.OUTPUTS
    None

.NOTES
    Author:            Martin Pugh
    Twitter:           @thesurlyadm1n
    Spiceworks:        Martin9700
    Blog:              www.thesurlyadmin.com
       
    Changelog:
       MLP - 2/1/18    Rename to New-VMwareSnapShot.  Updated to use VMware module, better credential handling 
                       and updated row coloring to group off username.  
       MLP             Fixed credential function (was calling it wrong, how come no one told me?!).  Added
                       PathToCredentials and Path to default to script path.  BIG: added new field
                       'Creator'.  Modified error handling (minor).  
       MLP             Updated Get-Credentials function to support domain level credentials in the
                       domainname\username format.
       MLP             By request added a calculation on how old the snapshot is in days. Discovered 
                       a "bug" when running the script on a VMware 4.1 system: the
                       SizeGB property does not exist!  Changed to use SizeMB and then manually
                       calculate the snapshot size in GB. Added some better error trapping.  Also
                       parameterized the email settings.
       MLP             Initial Release
.LINK
    https://github.com/martin9700/New-VMwareSnapShotReport
#>
[CmdletBinding()]
Param (
    [Alias("Host")]
    [string]$VIServer = "vcenterserver",
    [PSCredential]$Credential,
    [string]$Path,
    
    [string]$To = "martin@pughspace.com",
    [string]$From = "no-reply@vmwaresnapshot.com",
    [string]$SMTPServer = "yoursmtpserver"
)


#region Functions
Function Set-GroupRowColorsByColumn {
    <#
    #>
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory,ValueFromPipeline)]
        [string[]]$InputObject,
        [Parameter(Mandatory)]
        [string]$ColumnName,
        [string]$CSSEvenClass = "TREven",
        [string]$CSSOddClass = "TROdd"
    )
    Process {
        $NewHTML = ForEach ($Line in $InputObject)
        {
            If ($Line -like "*<th>*")
            {
                If ($Line -notlike "*$ColumnName*")
                {
                    Write-Error "Unable to locate a column named $ColumnName" -ErrorAction Stop
                }
                $Search = $Line | Select-String -Pattern "<th>.*?</th>" -AllMatches
                $Index = 0
                ForEach ($Column in $Search.Matches)
                {
                    If (($Column.Groups.Value -replace "<th>|</th>","") -eq $ColumnName)
                    {
                        Break
                    }
                    $Index ++
                }
            }
            If ($Line -like "*<td>*")
            {
                $Search = $Line | Select-String -Pattern "<td>.*?</td>" -AllMatches
                If ($LastColumn -ne $Search.Matches[$Index].Value)
                {
                    If ($Class -eq $CSSEvenClass)
                    {
                        $Class = $CSSOddClass
                    }
                    Else
                    {
                        $Class = $CSSEvenClass
                    }
                }
                $LastColumn = $Search.Matches[$Index].Value
                $Line = $Line.Replace("<tr>","<tr class=""$Class"">")
            }
            Write-Output $Line
        }
        Write-Output $NewHTML
    }
}


Function Get-SnapshotCreator {
    Param (
        [string]$VM,
        [datetime]$Created
    )

    (Get-VIEvent -Entity $VM -Types Info -Start $Created.AddSeconds(-10) -Finish $Created.AddSeconds(10) | Where FullFormattedMessage -eq "Task: Create virtual machine snapshot" | Select -ExpandProperty UserName).Split("\")[-1]
}
#endregion

If (-not (Get-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue))
{   Try { 
        Import-Module VMware.VimAutomation.Core -ErrorAction Stop 
    }
    Catch { 
        Write-Error "Problem loading VMware.VimAutomation.Core snapin because ""$_""" -ErrorAction Stop
    }
}

If (-not $Path)
{
    $Path = Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath Reports
    If (-not (Test-Path $Path))
    {
        $null = New-Item -Path $Path -ItemType Directory
    }
}
ElseIf (-not (Test-Path $Path))
{
    Write-Error "Output path ($Path) does not exist" -ErrorAction Stop
}

If (-not $Credential)
{
    $CredentialPath = Join-Path -Path (Split-Path $MyInvocation.MyCommand.Path) -ChildPath "Vcenter.xml"
    If (Test-Path -Path $CredentialPath)
    {
        $Credential = Import-Clixml -Path $CredentialPath
    }
    Else
    {
        $Credential = Get-Credential -Message "Please enter your VCenter credentials.  Make sure you are logged in under the service account that will be running this script"
        If (-not $Credential)
        {
            Write-Error "You did not enter credentials, aborting" -ErrorAction Stop
        }
        $Credential | Export-Clixml -Path $CredentialPath
    }
}

Write-Verbose "Connecting to $VIServer..."
Try {
    $Conn = Connect-VIServer $VIServer -Credential $Credential -ErrorAction Stop 3>$null
}
Catch {
    Write-Error "Error connecting to $VIServer because ""$_""" -ErrorAction Stop
}
Write-Verbose "Connected"

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TR:Hover TD {Background-Color: #C1D5F8;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
Snapshot Report - $VIServer
</title>
"@

Write-Verbose "Gathering snapshots (this will take awhile)..."
$Report = Get-VM | 
    Get-Snapshot | 
    Select VM,
    Name,
    Description,
    @{Name="SizeGB";Expression={ [math]::Round($_.SizeGB,2) }},
    @{Name="Creator";Expression={ Get-SnapshotCreator -VM $_.VM -Created $_.Created }},
    Created,
    @{Name="Days Old";Expression={ (New-TimeSpan -End (Get-Date) -Start $_.Created).Days }}

If (-not $Report)
{   $Report = [PSCustomObject]@{
        VM = "No snapshots found on any VM's controlled by $VIServer"
        Name = ""
        Description = ""
        SizeGB = ""
        Creator = ""
        Created = ""
        'Days Old' = ""
    }
}

Write-Verbose "Creating report and emailing"
$Report = $Report | 
    Sort Creator,VM | 
    ConvertTo-Html -Head $Header -PreContent "<p><h2>Snapshot Report - $VIServer</h2></p><br>" -PostContent "<p><br/><br/><h5>Run Date: $(Get-Date)</h5></p>" | 
    Set-GroupRowColorsByColumn -ColumnName Creator -CSSEvenClass even -CSSOddClass odd
$ReportName = Join-Path -Path $Path -ChildPath "SnapshotReport-$(Get-Date -Format 'MM-dd-yyyy').html"
$Report | Out-File -FilePath $ReportName


$MailSplat = @{
    To         = $To
    From       = $From
    Subject    = "$VIServer Snapshot Report"
    Body       = ($Report | Out-String)
    BodyAsHTML = $true
    SMTPServer = $SMTPServer
}

Send-MailMessage @MailSplat

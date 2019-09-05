<#
.SYNOPSIS Get/Install Windows Updates
.DESCRIPTION Based off the win32api https://docs.microsoft.com/en-us/windows/win32/wua_sdk/using-the-windows-update-agent-api
.Example Get-WindowsUpdates
.Example Install-WindowsUpdates
.AUTHOR Jason Connell
#>


if (-not ($PSVersionTable)) {Write-Warning 'PS1 Detected. PowerShell Version 2.0 or higher is required.';return}
if (-not ($PSVersionTable) -or $PSVersionTable.PSVersion.Major -lt 3 ) {Write-Verbose 'PS2 Detected. PowerShell Version 3.0 or higher may be required for full functionality.'}

$ModuleVersion = "1.0"

If ($env:PROCESSOR_ARCHITEW6432 -match '64' -and [IntPtr]::Size -ne 8) {
    Write-Warning '32-bit PowerShell session detected on 64-bit OS. Attempting to launch 64-Bit session to process commands.'
    $pshell="${env:WINDIR}\sysnative\windowspowershell\v1.0\powershell.exe"
    If (!(Test-Path -Path $pshell)) {
        Write-Warning 'SYSNATIVE PATH REDIRECTION IS NOT AVAILABLE. Attempting to access 64-bit PowerShell directly.'
        $pshell="${env:WINDIR}\System32\WindowsPowershell\v1.0\powershell.exe"
        $FSRedirection=$True
        Add-Type -Debug:$False -Name Wow64 -Namespace "Kernel32" -MemberDefinition @"
[DllImport("kernel32.dll", SetLastError=true)]
public static extern bool Wow64DisableWow64FsRedirection(ref IntPtr ptr);

[DllImport("kernel32.dll", SetLastError=true)]
public static extern bool Wow64RevertWow64FsRedirection(ref IntPtr ptr);
"@
        [ref]$ptr = New-Object System.IntPtr
        $Result = [Kernel32.Wow64]::Wow64DisableWow64FsRedirection($ptr) # Now you can call 64-bit Powershell from system32
    }
    If ($myInvocation.Line) {
        &"$pshell" -NonInteractive -NoProfile $myInvocation.Line
    } Elseif ($myInvocation.InvocationName) {
        &"$pshell" -NonInteractive -NoProfile -File "$($myInvocation.InvocationName)" $args
    } Else {
        &"$pshell" -NonInteractive -NoProfile $myInvocation.MyCommand
    }
    $ExitResult=$LASTEXITCODE
    If ($FSRedirection -eq $True) {
        [ref]$defaultptr = New-Object System.IntPtr
        $Result = [Kernel32.Wow64]::Wow64RevertWow64FsRedirection($defaultptr)
    }
    Write-Warning 'Exiting 64-bit session. Module will only remain loaded in native 64-bit PowerShell environment.'
Exit $ExitResult
}




function Get-WindowsUpdates {
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software'")
    return $SearchResults.Updates | select Title, MsrcSeverity
}


# Collects all updates and exports them to a CSV file for review
Function Get-WindowsUpdateToCSV {
<#
.SYNOPSIS
    This Function Will get a list of all availible patches

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    $SearchResults.Updates | select Title, SupportUrl, EulaAccepted, IsDownloaded, IsHidden, LastDeploymentChangeTime, MaxDownloadSize, MsrcSeverity | Export-Csv -Path C:\windows\Temp\WindowsUpdateExport.csv
}


# Install All Critical Patches
Function Install-AllCriticalPatches {

<#
.SYNOPSIS
    This Function will install all availible Critical updates

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>

    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

    $downloads = New-Object -ComObject Microsoft.Update.UpdateColl

    $count = 0
    foreach ($update in $SearchResults.Updates) {
        $Severity = $update | select MsrcSeverity
        $count = $count + 1

        if ($update."MsrcSeverity" -eq 'Critical') {
        Write-Output "Critical Pach Name: $($update."Title")"
        $downloads.Add($update)
        } Else {
        Write-Output "Not Critical Patch Name: $($update."Title")"
        }
    }

    # Download each update
    $downloader = $UpdateSession.CreateUpdateDownloader()
    $downloader.Updates = $downloads
    $downloader.Download()

    # Get list of downloaded updates so installs can happen.
    $installs = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($update in $downloads){
        if ($update.IsDownloaded){
            $installs.Add($update)
        }
    }


    $installer = $UpdateSession.CreateUpdateInstaller()
    $installer.Updates = $installs
    $installresult = $installer.Install()
    $installresult
}


# Install All Security Patches
Function Install-AllSecurityPatches {
<#
.SYNOPSIS
    This Function will install all security patches

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$SearchResults = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

$downloads = New-Object -ComObject Microsoft.Update.UpdateColl

$count = 0
foreach ($update in $SearchResults.Updates) {
    $Severity = $update | select MsrcSeverity
    $count = $count + 1

    if ($update."MsrcSeverity" -eq 'Security') {
    Write-Output "Security Pach Name: $($update."Title")"
    $downloads.Add($update)
    } Else {
    Write-Output "Not Security Patch Name: $($update."Title")"
    }
}

# Download each update
$downloader = $UpdateSession.CreateUpdateDownloader()
$downloader.Updates = $downloads
$downloader.Download()

# Get list of downloaded updates so installs can happen.
$installs = New-Object -ComObject Microsoft.Update.UpdateColl
foreach ($update in $downloads){
    if ($update.IsDownloaded){
        $installs.Add($update)
    }
}


$installer = $UpdateSession.CreateUpdateInstaller()
$installer.Updates = $installs
$installresult = $installer.Install()
$installresult
}




# Install All Patches
Function Install-AllPatches {

<#
.SYNOPSIS
    This Function Will install all patches

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
    $session = New-Object -ComObject Microsoft.Update.Session
    $searcher = $session.CreateUpdateSearcher()

    $result = $searcher.Search("IsInstalled=0 and Type='Software' and ISHidden=0")

    if ($result.Updates.Count -eq 0) {
         Write-Host "No updates to install"
    }
    else {
        $result.Updates | select Title
    }

    $downloads = New-Object -ComObject Microsoft.Update.UpdateColl

    foreach ($update in $result.Updates){
         $downloads.Add($update)
    }

    $downloader = $session.CreateUpdateDownLoader()
    $downloader.Updates = $downloads
    $downloader.Download()

    $installs = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($update in $result.Updates){
         if ($update.IsDownloaded){
               $installs.Add($update)
         }
    }

    $installer = $session.CreateUpdateInstaller()
    $installer.Updates = $installs
    $installresult = $installer.Install()
    $installresult
}


# Open connection to the Windows Update API. Return Results
Function Test-WindowsUpdateConnection {
<#
.SYNOPSIS
    This Function will open a test connection for windows update

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
}


# Gets a list of all installed Updates
Function Get-InstalledUpdates {

<#
.SYNOPSIS
    This Function will return a list of installed windows updates

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019

#>
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResults = $UpdateSearcher.Search("IsInstalled=1 and Type='Software'")
    $result = $SearchResults.Updates | select Title, LastDeplymentChangeTime
    Write-Output $result
}





Function Stop-WindowsUpdateServices($service) {

<#
.SYNOPSIS
    This Function will stop a given service.

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
    Try{

        if ((Get-Service -Name $service).Status -eq 'Stopped') {
            return
        }
        Else {


            # Stop Windows Upadate service
            Try{
                Stop-Service -Name $service -Force
            }
            Catch {
                Write-Output 'Service Failed to Stop'
            }
        }
    }
    Catch{
        Write-Output 'Failed to Stop Service'
    }
}


Function Start-WindowsUpdateServices($service) {

<#
.SYNOPSIS
    This Function will start a given service

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>
    Try{
        if ((Get-Service -Name $service).Status -eq 'Stopped'){
            Start-Service -Name $service
        }
        Else{
            return
        }
    }
    Catch {
        Write-Output 'Failed to Start Service'
    }
}


Function Reset-SoftwareDistribution {
<#
.SYNOPSIS
    This Function will rebuild the softwaredistribution folder

.NOTES
    Version: 1.0
    Author: Jason Connell
    Creation Date: 8/09/2019
#>

    $WIN_DIR = (get-childitem env:SystemRoot).Value
    $SOFTWAREDISTRIBUTION_DIR = Join-Path -Path $WIN_DIR -ChildPath "\SoftwareDistribution"
    $Services = @('wuauserv', 'BITS', 'CryptSVC')

    # Stop Services
    foreach ($item in $Services){
        Stop-WindowsUpdateServices($item)
    }

    # Get Current datetime to append to filename
    $curentDateTime = Get-Date -Format "MM:dd:yyyy:HH:mm:ss" | ForEach-Object {$_ -replace":", "-" }
    $NewSoftwareDistributionPath = $SOFTWAREDISTRIBUTION_DIR + $curentDateTime


    # Rename SoftwareDistribution Folder
    Try{
      Rename-Item $SOFTWAREDISTRIBUTION_DIR -NewName $NewSoftwareDistributionPath
    }
    Catch{
      Write-Output "Failed to Rename  SoftwareDistribution Folder"
      break
    }


    # Start Services
    foreach ($item in $Services){
        Start-WindowsUpdateServices($item)
    }
    Write-Output 'SoftwareDistribution Folder Has Been Rebuilt'
}


Function Schedule-PatchInstall {
  <#
  .SYNOPSIS
      This Function will add a scheduled task to reboot the computer.

  .Example
    Schedule-PatchInstall -Date '08/14/2019' -Time '6:00pm' -ScheduleName 'Reboot For Patch Install'
  .NOTES
      Version: 1.0
      Author: Jason Connell
      Creation Date: 8/14/2019
  #>

    param (
        [string]$Time = '',
        [string]$Date = '',
        [string]$ScheduleName = ''
    )


    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $adminStatus = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-NOT($adminStatus)){
        Write-Output 'This procedure needs to be run as administrator'
        break
        }

    if ($Time -eq '') {
        Write-Output 'Must specify time with -Time argument'
        break
        }


    if ($Date -eq ''){
        Write-Output 'Must specify date with -Date argument'
        break
        }


    if ($ScheduleName -eq ''){
        Write-Output 'Must specify schedule name with -ScheduleName argument'
        break
        }



    # Set up scheduled task
    $action = New-ScheduledTaskAction -Execute 'Powershell.exe' -Argument 'Restart-Computer -Force'

    $trigger =  New-ScheduledTaskTrigger -Once -At "$($Date) $($Time)"

    Register-ScheduledTask -Action $action -Trigger $trigger -TaskName $ScheduleName -Description "Scheduled Reboot for patch install"
}






$PublicFunctions=@(((@"
Get-WindowsUpdates
Get-WindowsUpdateToCSV
Install-AllCriticalPatches
Install-AllSecurityPatches
Install-AllPatches
Test-WindowsUpdateConnection
Get-InstalledUpdates
Stop-WindowsUpdateServices
Start-WindowsUpdateServices
Reset-SoftwareDistribution
Schedule-PatchInstall
"@) -replace "[`r`n,\s]+",',') -split ',')


Function Get-ListOfFunctions {
  foreach ($func in $PublicFunctions) {
    Write-Output $func
  }
}

# ps-server-cleaner.ps1
# Cleans PennyMac servers before and after patch night
# Contact Elliott Berglund x8981 if you have any issues
Clear-Host
$Host.UI.RawUI.BufferSize.Height = 2000

# Make some blank lines to prevent overlap with progress bar at the beginning of the run
"`n`n`n`n`n"

$LogPath = "$PSScriptRoot\Logs"

# Make log directory if it doesn't exist
Try
{
    If (!([System.IO.Directory]::Exists($LogPath)))
    {
        New-Item -ItemType Directory -Path "$PSScriptRoot" -Name "Logs" -ErrorAction Stop
    }
    Else
    {
        
    }
}
Catch
{
    Write-Host "Unhandled exception creating $LogPath" -ForegroundColor Yellow
    Write-Host "Exception: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
}
Finally
{
    Start-Transcript -OutputDirectory $LogPath
}

# Configuration options
$ServerListCSV = "serverlist.csv" # Filename of CSV listing all servers to be cleaned, should be placed in the same directory as this script
$CSVHeader = "HostName" # Header of the column in the CSV file containing the machines DNS name or IP address
$VerbosePreference = "SilentlyContinue" # Toggle Verbosity, "SilentlyContinue" to suppress VERBOSE messages, "Continue" to use full Verbosity

# Non user specific cleanup targets, relative from C:
$CleanupTargets = New-Object System.Collections.ArrayList
$CleanupTargets += '$RECYCLE.BIN'
$CleanupTargets += "Windows\Temp"
$CleanupTargets += "Temp"
$CleanupTargets += "Windows\ProPatches\Patches"
$CleanupTargets += "Windows\MiniDump"
$CleanupTargets += "Windows\LiveKernelReports"
$CleanupTargets += "NVIDIA\DisplayDriver"

# User specifuc cleanup targets, relative from user profile directory (C:\Users\{username}\)
$UserCleanupTargets = New-Object System.Collections.ArrayList
$UserCleanupTargets += "AppData\Local\Temp"
$UserCleanupTargets += "AppData\Local\Microsoft\Windows\INetCache"
$UserCleanupTargets += "AppData\Local\Microsoft\Windows\Temporary Internet Files"
$UserCleanupTargets += "AppData\Local\Google\Chrome\User Data\Default\Cache"
$UserCleanupTargets += "AppData\Local\Google\Chrome\Update"
$UserCleanupTargets += "AppData\Local\Google\Chrome SxS\User Data\Default\Cache"
$UserCleanupTargets += "AppData\Local\Google\Chrome SxS\Update"
$UserCleanupTargets += "AppData\Local\CrashDumps"
$UserCleanupTargets += "AppData\Local\Microsoft\Terminal Server Client\Cache"
$UserCleanupTargets += "AppData\LocalLow\Sun\Java\Deployment\cache\6.0"

# Import AD Module
If (!(Get-Module -Name ActiveDirectory))
{
    Import-Module -Name ActiveDirectory -ErrorAction Stop
}

# FUNCTIONS START
Function Test-PathEx
{
    Param($Path)

    If (Test-Path -LiteralPath $Path)
    {
        $True
    }
    Else
    {
        [System.IO.Directory]::Exists($Path)
    }
}

Function Get-FreeSpace
{
    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact="Low")]

    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,

        [Parameter(Mandatory=$True, Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]$DriveLetter
)
    Begin
    {
        $DriveLetter = $DriveLetter -replace '[:|$]',''
    }
    Process
    {
        $FreeSpace = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $ComputerName | Where-Object { $_.DeviceID -eq "$DriveLetter`:" } | Select-Object SystemName, Caption, @{Name="FreeSpace"; Expression={"$([math]::Round($_.FreeSpace / 1GB,2))GB"}}, @{Name="PercentFree"; Expression={"$([math]::Round($_.FreeSpace / $_.Size,2) * 100)%"}}
        $Out = [PSCustomObject][Ordered]@{
        'ComputerName' = $FreeSpace.SystemName
        'DriveLetter' = $FreeSpace.Caption
        'FreeSpace' = $FreeSpace.FreeSpace
        'PercentFree' = $FreeSpace.PercentFree
        }
    }
    End
    {
        Return $Out
    }
}

Function Resolve-Host
{
    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact="Low")]

    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName
    )

    Begin
    {

    }

    Process
    {
        If ($ComputerName -As [IPAddress])
        {
            $IP = $ComputerName
            $ComputerName = [System.Net.Dns]::GetHostEntry($ComputerName).HostName
        }
        Else
        {
            $ComputerName = $ComputerName.ToUpper()
            $IP = [System.Net.Dns]::GetHostAddresses($ComputerName).IPAddressToString
        }
    }

    End
    {
        Return [PSCustomObject][Ordered]@{
        'ComputerName' = $ComputerName
        'IPAddress' = $IP
        }
    }
}

Function Remove-WithProgress
{
    [CmdletBinding(SupportsShouldProcess=$True, ConfirmImpact="Medium")]

    Param(
        [Parameter(Mandatory=$True, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$ComputerName,

        [Parameter(Mandatory=$True, Position=2)]
        [ValidateNotNullOrEmpty()]
        [String]$DriveLetter,

        [Parameter(Mandatory=$True, Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]$Path
    )

    Begin
    {
        # Set Variables
        $CurrentFileCount = 0
        $CurrentFolderCount = 0
        $DriveLetter = $DriveLetter -replace '[:|$]',''
        $CombinedPath = Join-Path -Path "\\$ComputerName" -ChildPath "$DriveLetter$" | Join-Path -ChildPath "$Path"

        # Start progress bar
        Write-Progress -Id 1 -Activity "Enumerating files on $ComputerName" -PercentComplete 0

        # Output to screen
        Write-Host $("`n" + ('=' * 60) + "`n") -ForegroundColor Cyan
        Write-Host "Path: $CombinedPath"

        # Check if path is writable
        Try
        {
            New-Item -Path $CombinedPath -Name "writetest.tmp" -ItemType File -Force -ErrorAction Stop | Remove-Item | Out-Null
            
        }
        Catch
        {
            Write-Host "Cannot write to $CombinedPath" -ForegroundColor Red
            Continue
        }
    }

    Process
    {
        # Enumerate files
        Try
		{
			$Files = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -Attributes !D,!D+H,!D+S,!D+H+S -Recurse -ErrorAction Stop) | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
		}
		Catch
		{
            Write-Host "Unhandled exception enumerating $CombinedPath"
			Write-Host $_.Exception.Message -ForegroundColor Red
		}

        # Total file count for progress bar
        $FileCount = ($Files | Measure-Object).Count
        $TotalSize = [Math]::Round(($Files | Measure-Object -Sum -Property Length).Sum /1GB,3)

        "Removing $FileCount files... $TotalSize`GB."

        ForEach ($File in $Files)
        {
            $CurrentFileCount++
            $Percentage = [math]::Round(($CurrentFileCount / $FileCount) * 100)
            Write-Progress -Id 1 -Activity "Removing Files" -CurrentOperation "File: $($File.FullName)" -PercentComplete $Percentage -Status "Progress: $CurrentFileCount of $FileCount, $Percentage%"
            Try
            {
                Write-Verbose -Message "Removing file $($File.FullName)"
                $File | Remove-Item -ErrorAction Stop
            }
            Catch [System.IO.IOException]
            {
                Write-Host "$($_.Exception.Message) while deleting $($File.FullName)" -ForegroundColor Red
            }
            Catch [System.UnauthorizedAccessException]
            {
                Write-Host "$($_.Exception.Message) while deleting $($File.FullName)" -ForegroundColor Red
            }
            Catch
            {
                Write-Host "$($_.Exception.Message) while deleting $($File.FullName)" -ForegroundColor Red
                Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
            }
            Finally
            {
                
            }
        }

        # Reset progress bar for phase 2
        Write-Progress -Id 1 -Activity "Enumerating empty directories on $ComputerName" -CurrentOperation "Path: $CombinedPath" -PercentComplete 0

        # Enumerate folders with 0 files
        Try
		{
			$EmptyFolders = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -Attributes D,D+H,D+S,D+S+H -Recurse -ErrorAction Stop) | Where-Object {($_.GetFiles()).Count -eq 0} -ErrorAction Stop | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
		}
		Catch
		{
            Write-Host "Unhandled exception enumerating $CombinedPath" -ForegroundColor Yellow
			Write-Host $_.Exception.Message -ForegroundColor Red
		}
    
        # How many empty folders for progress bars
        $EmptyCount = ($EmptyFolders | Measure-Object).Count

        If ($EmptyCount -gt 0)
        {
            ForEach ($EmptyFolder in $EmptyFolders)
            {
                # Increment Folder Counter
                $CurrentFolderCount++

                $Percentage = [Math]::Round(($CurrentFolderCount / $EmptyCount) * 100)
        
                If ((($EmptyFolder.GetFiles()).Count + ($EmptyFolder.GetDirectories()).Count) -ne 0)
                {
                    Write-Verbose -Message "$($EmptyFolder.FullName) not empty, skipping..."
                    Continue
                }
                Write-Progress -Id 1 -Activity "Removing Empty Directories" -CurrentOperation "Removing Empty Directory: $($EmptyFolder.FullName)" -PercentComplete "$Percentage" -Status "Progress: $CurrentFolderCount of $EmptyCount, $Percentage%"
                Try
                {
                    Write-Verbose -Message "Removing folder $($EmptyFolder.FullName)"
                    $EmptyFolder | Remove-Item -ErrorAction Stop
                }
                Catch [System.IO.IOException]
                {
                    Write-Host "$($_.Exception.Message) while deleting $($EmptyFolder.FullName)" -ForegroundColor Red
                }
                Catch [System.UnauthorizedAccessException]
                {
                    Write-Host "$($_.Exception.Message) while deleting $($EmptyFolder.FullName)" -ForegroundColor Red
                }
                Catch
                {
                    Write-Host "$($_.Exception.Message) while deleting $($EmptyFolder.FullName)" -ForegroundColor Red
                    Write-Host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
                }
                Finally
                {
                    
                }
            }
        }
    }

    End
    {
        # Close progress bar
        Write-Progress -Id 1 -Completed -Activity 'Done'
        Return
    }
}
# END FUNCTIONS

# Prepare array and import CSV
$ServerList = @()
Write-Host "Importing $ServerListCSV"
Try
{
    # Import CSV of hostnames then remove any characters invalid in computer names
    Import-CSV -LiteralPath "$PSScriptRoot\$ServerListCSV" | ForEach-Object {$ServerList += $_."$CSVHeader" -replace "[`r|`n|`t|`:|`*|`\|`/|`?|`"|`<|`>|`||`,|`~|`!|`^|`@|`#|`%|`|`'|`&|`.|`_|`(|`)|`{|`}| ]",""}
}
Catch
{
    Write-Host "$($_.Exception.Message)"
    "No $ServerListCSV file found in $PSScriptRoot. Exiting."
    Exit
}
Write-Host "Done"

# Progress Bar stuff
$TotalServers = ($ServerList | Measure-Object).Count
"Machines Imported: $TotalServers"
$Counter = 0

Write-Progress -Activity "Recovering Disk Space on Servers" -CurrentOperation "Starting" -Id 0 -PercentComplete -1 -Status "Processing"
ForEach ($Server in $ServerList)
{
    # Set drive letter for cleanup
    $DriveLetter = "C"

    # Create UserProfiles, DeleteMethodList,and FailureList arrays
    $UserProfiles = @()
    $DeleteMethodList = @()

    # Progress bar
    $Counter++
    $Percentage = [Math]::Round(($Counter / $TotalServers) * 100)

    If ($Resolved)
    {
        Remove-Variable -Name Resolved
    }
    If ($ServerHostName)
    {
        Remove-Variable -Name ServerHostName
    }
    If ($ServerIP)
    {
        Remove-Variable -Name ServerIP
    }

    $Resolved = Resolve-Host -ComputerName $Server
    $ServerHostName = $Resolved.ComputerName
    $ServerIP = $Resolved.IPAddress

    If (!(Test-NetConnection -ComputerName $ServerHostName -InformationLevel Quiet))
    {
        Write-Warning -Message "$Server did not respond, skipping"
        Continue
    }

    If ([System.IO.Directory]::Exists("\\$ServerHostName\c$") -eq $False)
    {
        Write-Warning -Message "$Server does not respond to UNC file requests. Skipping."
        Continue
    }

    Write-Progress -Activity "Recovering Disk Space on Servers" -CurrentOperation "Server Name: $Server" -Id 0 -PercentComplete $Percentage -Status "$Counter of $TotalServers, $Percentage%"

    # Free space
    Write-Host $("`n" + ('=' * 60) + "`n") -ForegroundColor Green
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter $DriveLetter | Format-Table
    Write-Host $(('=' * 60) + "`n") -ForegroundColor Green

    # Get profile list excluding local and service accounts, the local administrator account, and IIS App Pools
    $Blacklist = @("Administrator",".NET v2.0 Classic",".NET v4.5 Classic",".NET v2.0",".NET v4.5","Classic .NET AppPool","DefaultAppPool")
    $LocalAccounts = @(Get-WmiObject -Class Win32_UserAccount -ComputerName $ServerHostName -Filter "LocalAccount='True'" | Select-Object -ExpandProperty SID)
    ForEach ($SA in @(Get-WmiObject -Class Win32_Service -ComputerName $ServerHostName | Select-Object -ExpandProperty StartName))
    {
        $Blacklist += $SA.Split("\")[-1]
    }
    $UserProfiles = @(Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "Special='False'" | Where-Object {($_.SID -notin $LocalAccounts) -and ($_.LocalPath.Split("\")[-1] -notin $Blacklist) -and ($_.LocalPath.Split("\")[-1] -notlike "00*") -and ($_.LocalPath.Split("\")[-1] -notlike "*MSSQL*") -and ($_.LocalPath.Split("\")[-1] -notlike "*MsDts*")})

    ForEach ($UserProfile in $UserProfiles)
    {
        $UserLocalPath = $UserProfile.LocalPath.Split("\")[-1]
        If ($SID)
        {
            Remove-Variable -Name SID
        }
        # If profile status Bit Field includes 8 (corrupt profile), flag for removal.
        Write-Host "Checking user profile: $UserLocalPath"

        If ((8 -band $UserProfile.Status) -eq 8)
        {
            Write-Warning -Message "PROFILE CORRUPT!"
            Write-Host "Flagged `"$UserLocalPath`" for removal." -ForegroundColor Yellow
            $DeleteMethodList += $UserProfile
            Continue
        }

        $SID = $UserProfile | Select-Object -ExpandProperty sid

        # Check against AD to see if user account exists
        Try
        {
            Get-ADUser -Identity $SID | Out-Null
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {
            Write-Host "$($_.Exception.Message)"
            Write-Warning -Message "Profile $UserLocalPath {SID: $SID} does not exist in Active Directory."
            Write-Host " > Flagged `"$UserLocalPath`" for removal." -ForegroundColor Red
            $DeleteMethodList += $UserProfile
            Continue
        }
        Catch
        {
            Write-Host "$($_.Exception.Message)"
            Write-Warning -Message "Unhandled Exception with Active Directory module, skipping $UserLocalPath."
            Continue
        }
        Write-Host " > $UserLocalPath - Status OK" -ForegroundColor Green
    }
    Write-Host "Scanned all profiles on $ServerHostName"
    "{0} of {1} local profiles scheduled for deletion on {2}" -F ($DeleteMethodList | Measure-Object).Count,($UserProfiles | Measure-Object).Count,$ServerHostName

    $RemovalCount = 0
    $FailureCount = 0
    ForEach ($User in $DeleteMethodList)
    {
        $UserLocalPath = $User.LocalPath.Split("\")[-1]
        Write-Host "Deleting Profile: $UserLocalPath" -ForegroundColor Yellow
        Try
        {
            $User.Delete()
        }
        Catch
        {
            Write-Host "An error occurred deleting $UserLocalPath!" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
            Continue
        }
        Write-Host "Success!"
        $RemovalCount++
    }
    If ($RemovalCount -gt 0)
    {
        Write-Host "$RemovalCount unused profile(s) removed successfully." -ForegroundColor Cyan
    }

    If ($FailureCount -gt 0)
    {
        Write-Host "$FailureCount profile(s) failed to delete."
    }

    ForEach ($p in $CleanupTargets)
    {
        $CleanTarget = $p
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $CleanTarget
    }

    $RemainingProfiles = @(Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName)
    ForEach ($UserProfile in $RemainingProfiles)
    {
        $BasePath = $UserProfile.LocalPath -replace '[C|c]:\\',''

        ForEach ($p in $UserCleanupTargets)
        {
            $CleanTarget = $BasePath + '\' + $p
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $CleanTarget
        }
    }
    Write-Host $("`n" + ('=' * 60) + "`n") -ForegroundColor Green
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter $DriveLetter | Format-Table
    Write-Host $(('=' * 60) + "`n") -ForegroundColor Green
    Write-Host "Cleanup completed on $ServerHostName"
    Write-Host $("`n" + ('=' * 60) + "`n") -ForegroundColor Cyan
}

Write-Progress -Id 0 -Completed -Activity 'Done'
Stop-Transcript
Write-Host "Job complete. Check $LogPath for log files"

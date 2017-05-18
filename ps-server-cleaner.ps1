# ps-server-cleaner.ps1
# Cleans PennyMac servers before and after patch night
# Contact Elliott Berglund x8981 if you have any issues
Clear-Host

# Make some blank lines to prevent overlap with progress bar at the beginning of the run
"`n`n`n`n`n"

# Configuration options
$ServerListCSV = "serverlist.csv" # Filename of CSV listing all servers to be cleaned, should be placed in the same directory as this script
$CSVHeader = "HostName" # Header of the column in the CSV file containing the machines DNS name or IP address
$VerbosePreference = "SilentlyContinue" # Toggle Verbosity, "SilentlyContinue" to suppress VERBOSE messages, "Continue" to use full Verbosity
$ErrorActionPreference = "Continue" # Toggle Error Output. "SilentlyContinue" suppresses errors, "Continue" shows all errors

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
        #$Parent = Split-Path -LiteralPath $Path
        #[System.IO.Directory]::EnumerateFiles($Parent) -Contains $Path
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
        $FreeSpace = Get-WmiObject -Class Win32_LogicalDisk -ComputerName $ComputerName |
        Where-Object { $_.DeviceID -eq "$DriveLetter`:" } |
        Select-Object @{Name="ComputerName"; Expression={ $_.SystemName } }, @{Name="DriveLetter"; Expression={ $_.Caption } }, @{Name="FreeSpace"; Expression={ "$([math]::Round($_.FreeSpace / 1GB,2))GB" } }, @{Name="PercentFree"; Expression={"$([math]::Round($_.FreeSpace / $_.Size,2) * 100)%"}} | Format-List
    }
    End
    {
        Return $FreeSpace
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
            $IP = [System.Net.Dns]::GetHostAddresses($ComputerName).IPAddressToString[-1]
        }
    }

    End
    {
        Return $ComputerName, $IP
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
        $DriveLetter = $DriveLetter -replace '[:|$]',''
        $CombinedPath = Join-Path -Path "\\$ComputerName" -ChildPath "$DriveLetter$" | Join-Path -ChildPath "$Path"

        # Start progress bar
        Write-Progress -Id 1 -Activity "Enumerating files on $ComputerName" -PercentComplete 0
    }

    Process
    {
        # Progress Bar counter
        $CurrentFileCount = 0
        $CurrentFolderCount = 0

        # Enumerate files, silence errors
        $Files = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -ErrorAction SilentlyContinue -Attributes !D,!D+H,!D+S,!D+H+S) | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
        # Timer Stop

        # Total file count for progress bar
        $FileCount = ($Files | Measure-Object).Count
        $TotalSize = ($Files | Measure-Object -Sum -Property Length).Sum
        $TotalSize = [math]::Round($TotalSize / 1GB,3)

        "Removing $FileCount files... $TotalSize`GB."

        If ($DeleteFailed)
        {
            Remove-Variable -Name DeleteFailed
        }
        $DeleteFailed = @()
        $DeleteFail = $False
        ForEach ($File in $Files)
        {
            $CurrentFileCount++
            $FullFileName = $File.FullName
            $Percentage = [math]::Round(($CurrentFileCount / $FileCount) * 100)
            Write-Progress -Id 1 -Activity "Removing Files" -CurrentOperation "File: $FullFileName" -PercentComplete $Percentage -Status "Progress: $CurrentFileCount of $FileCount, $Percentage%"
            Write-Verbose -Message "Removing file $FullFileName"
            Try
            {
                $File | Remove-Item -Force
            }
            Catch
            {
                Write-Host "$($Error[0].Exception.Message)"
                $DeleteFail = $True
                $DeleteFailed += "$FullFileName delete failed"
            }
        }

        # Reset progress bar for phase 2
        Write-Progress -Id 1 -Activity "Enumerating empty directories on $ComputerName" -CurrentOperation "Path: $CombinedPath" -PercentComplete 0

        # Enumerate remaining files
        $RemainingFiles = @(Get-ChildItem -Force -LiteralPath $CombinedPath -ErrorAction SilentlyContinue -Attributes !D,!D+H,!D+S,!D+H+S).Count
        # Enumerate folders with 0 files
        $EmptyFolders = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -Attributes D,D+H,D+S,D+S+H -ErrorAction SilentlyContinue) | Where-Object {($_.GetFiles()).Count -eq 0} | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
    
        # How many empty folders for progress bars
        $EmptyCount = ($EmptyFolders | Measure-Object).Count

        If ($EmptyCount -gt 0)
        {
            ForEach ($EmptyFolder in $EmptyFolders)
            {
                # Increment Folder Counter
                $CurrentFolderCount++

                # Full Folder Name
                $FullFolderName = $EmptyFolder.FullName

                $Percentage = [math]::Round(($CurrentFolderCount / $EmptyCount) * 100)
        
                If ((($EmptyFolder.GetFiles()).Count + ($EmptyFolder.GetDirectories()).Count) -ne 0)
                {
                    Write-Verbose -Message "$FullFolderName not empty, skipping..."
                    Continue
                }
                Write-Progress -Id 1 -Activity "Removing Empty Directories" -CurrentOperation "Removing Empty Directory: $FullFolderName" -PercentComplete "$Percentage" -Status "Progress: $CurrentFolderCount of $EmptyCount, $Percentage%"
                Write-Verbose -Message "Removing folder $FullFolderName"
                Try
                {
                    $EmptyFolder | Remove-Item -Force
                }
                Catch
                {
                    Write-Host "$($Error[0].Exception.Message)"
                    $DeleteFail = $True
                    $DeleteFailed += "$FullFolderName folder delete failed"
                }
            }
        }
    }

    End
    {
        # Close progress bar
        Write-Progress -Id 1 -Completed -Activity 'Done'
        If ($DeleteFail -eq $True)
        {
            $DeleteFailed | Out-File -LiteralPath "$LogPath\$ServerHostName-skippedfiles-$LogDate.log" -Append
        }
        Return
    }
}
# END FUNCTIONS

# Import AD Module
If (!(Get-Module -Name ActiveDirectory))
{
    Import-Module -Name ActiveDirectory -ErrorAction Stop
}

# Prepare logging
$LogDate = (Get-Date).ToString('yyyy-MM-dd')
# Make log directory if it doesn't exist
If ([System.IO.Directory]::Exists("$PSScriptRoot\logs\servercleanup") -eq $False)
{
    New-Item -Path "$PSScriptRoot\logs\servercleanup" -ItemType Directory
}
$LogPath = "$PSScriptRoot\logs\servercleanup"

# Prepare array and import CSV
$ServerList = @()
Write-Host -Object "Importing $ServerListCSV"
Try
{
    Import-CSV -LiteralPath "$PSScriptRoot\$ServerListCSV" | ForEach-Object {$ServerList += $_."$CSVHeader" -replace "`r`n","" -replace "`t","" -replace " ","" }
}
Catch
{
    Write-Host "$($Error[0].Exception.Message)"
    "No $ServerListCSV file found in $PSScriptRoot. Exiting."
    Exit
}
Write-Host -Object "Done"

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
    $FailureList = @()

    # Progress bar
    $Counter++
    $Percentage = [math]::Round(($Counter / $TotalServers) * 100)

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
    $ServerHostName = $Resolved[0]
    $ServerIP = $Resolved[-1]

    If (!(Test-NetConnection -ComputerName $ServerHostName -InformationLevel Quiet))
    {
        Write-Warning -Message "$Server did not respond, skipping" | Tee-Object -FilePath "$LogPath\skipped-$LogDate.log" -Append
        Continue
    }

    If ([System.IO.Directory]::Exists("\\$ServerHostName\c`$") -eq $False)
    {
        Write-Warning -Message "$Server does not respond to UNC file requests. Skipping." | Tee-Object -FilePath "$LogPath\skipped-$LogDate.log" -Append
        Continue
    }

    Write-Progress -Activity "Recovering Disk Space on Servers" -CurrentOperation "Server Name: $Server" -Id 0 -PercentComplete $Percentage -Status "$Counter of $TotalServers, $Percentage%"

    # Free space
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter $DriveLetter | Tee-Object -FilePath "$LogPath\$ServerHostName-freespace-$LogDate.log" -Append
    #Start-Sleep -Seconds 2

    # Get profile list excluding service accounts and the local administrator account
    $LocalAccounts = @(Get-WmiObject -Class Win32_UserAccount -ComputerName $ServerHostName -Filter "LocalAccount='True'")
    $ExcludedAccounts = @("MsDtsServer100","MsDtsServer110","MsDtsServer120","ReportServer","MSSQLFDLauncher","MSSQLSERVER","SQLSERVERAGENT","Administrator","launcher-v4",".NET v2.0 Classic",".NET v4.5 Classic",".NET v2.0",".NET v4.5","Classic .NET AppPool")
    $UserProfiles = @(Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "Special='False'" | Where-Object {$_.LocalPath.Split("\")[-1] -notin $LocalAccounts.Caption.Split("\") -and $_.LocalPath.Split("\")[-1] -notin $ExcludedAccounts})

    ForEach ($UserProfile in $UserProfiles)
    {
        $UserLocalPath = $UserProfile.LocalPath.Split("\")[-1]
        If ($SID)
        {
            Remove-Variable -Name SID
        }
        # If profile status Bit Field includes 8 (corrupt profile), flag for removal.
        Write-Host -Object "Checking user profile: $UserLocalPath"

        If ((8 -band $UserProfile.Status) -eq 8)
        {
            Write-Warning -Message "PROFILE CORRUPT!"
            Write-Host -Object "Flagged `"$UserLocalPath`" for removal." -ForegroundColor Yellow
            $DeleteMethodList += $UserProfile
            Continue
        }

        $SID = $UserProfile | Select-Object -ExpandProperty sid

        #Check against AD to see if user account exists
        Try
        {
            Get-ADUser -Identity $SID | Out-Null
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {
            Write-Host "$($Error[0].Exception.Message)"
            Write-Warning -Message "Profile $UserLocalPath {SID: $SID} does not exist in Active Directory."
            Write-Host -Object " > Flagged `"$UserLocalPath`" for removal." -ForegroundColor Red
            $DeleteMethodList += $UserProfile
            Continue
        }
        Catch
        {
            Write-Host "$($Error[0].Exception.Message)"
            Write-Warning -Message "Unhandled Exception with Active Directory module, skipping $UserLocalPath."
            Continue
        }
        Write-Host -Object " > $UserLocalPath - Status OK" -ForegroundColor Green
    }
    Write-Host -Object "Scanned all profiles on $ServerHostName"
    "{0} of {1} local profiles scheduled for deletion on {2}" -F ($DeleteMethodList | Measure-Object).Count,($UserProfiles | Measure-Object).Count,$ServerHostName

    $RemovalCount = 0
    $FailureCount = 0
    ForEach ($User in $DeleteMethodList)
    {
        $UserLocalPath = $User.LocalPath.Split("\")[-1]
        Write-Host -Object "Deleting Profile: $UserLocalPath" -ForegroundColor Yellow
        Try
        {
            $User.Delete()
        }
        Catch
        {
            Write-Host -Object "An error occurred deleting $UserLocalPath!" -ForegroundColor Red
            Write-Host -Object  $Error[0]
            $FailureList += $User.LocalPath
            $FailureCount++
            Continue
        }
        Write-Host -Object "Success!"
        $RemovalCount++
    }
    If ($RemovalCount -gt 0)
    {
        Write-Host -Object "$RemovalCount unused profile(s) removed successfully." -ForegroundColor Cyan
    }

    If ($FailureCount -gt 0)
    {
        Write-Host -Object "$FailureCount profile(s) failed to delete. Check `"$LogPath\$ServerHostName-userprofiles-$LogDate.log`" for details."
        $FailureList | Out-File -LiteralPath "$LogPath\$ServerHostName-userprofiles-$LogDate.log"
    }

    # C:\$RECYCLE.BIN
    $RelativePath = '$RECYCLE.BIN'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If ([System.IO.Directory]::Exists("$PathTest"))
    {
        Write-Host -Object "Path: $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\Temp
    $RelativePath = 'Windows\Temp'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If ([System.IO.Directory]::Exists("$PathTest"))
    {
        Write-Host -Object "Path: $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Temp
    $RelativePath = 'Temp'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If ([System.IO.Directory]::Exists("$PathTest"))
    {
        Write-Host -Object "Path: $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\ProPatches\Patches
    $RelativePath = 'Windows\ProPatches\Patches'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If ([System.IO.Directory]::Exists("$PathTest"))
    {
        Write-Host -Object "Path: $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    $RemainingProfiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "Special='False'" | Where-Object {$_.LocalPath.Split("\")[-1] -notcontains $LocalAccounts}

    ForEach ($UserProfile in $RemainingProfiles)
    {
        $UserPath = $UserProfile.LocalPath
        $UserPath = $UserPath -replace '[C|c]:\\',''

        # User temp files
        $RelativePath = "$UserPath\AppData\Local\Temp"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User IE Cache (New)
        $RelativePath = "$UserPath\AppData\Local\Microsoft\Windows\INetCache"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User IE Cache (Old)
        $RelativePath = "$UserPath\AppData\Local\Microsoft\Windows\Temporary Internet Files"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Chrome Cache
        $RelativePath = "$UserPath\AppData\Local\Google\Chrome\User Data\Default\Cache"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Chrome Updates
        $RelativePath = "$UserPath\AppData\Local\Google\Chrome\Update"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Crash Dumps
        $RelativePath = "$UserPath\AppData\Local\CrashDumps"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If ([System.IO.Directory]::Exists("$PathTest"))
        {
            Write-Host -Object "Path: $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
    }
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter "$DriveLetter" | Tee-Object -FilePath "$LogPath\$ServerHostName-freespace-$LogDate.log" -Append
    Write-Host -Object "Cleanup completed on $ServerHostName"
}
Write-Progress -Id 0 -Completed -Activity 'Done'
Write-Host -Object "Job complete. Check $LogPath for log files"

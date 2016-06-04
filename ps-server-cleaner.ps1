#requires -Version 3.0
#requires -RunAsAdministrator
# ps-server-cleaner.ps1
# Cleans PennyMac servers before and after patch night
# Contact Elliott Berglund x8981 if you have any issues
Clear-Host

# Configuration options
$ServerListCSV = "serverlist.csv" # Filename of CSV listing all servers to be cleaned, should be placed in the same directory as this script
$CSVHeader = "Hostname" # Header of the column in the CSV file containing the machines DNS name or IP address
$VerbosePreference = "SilentlyContinue" # Toggle Verbosity, "SilentlyContinue" to suppress VERBOSE messages, "Continue" to use full Verbosity
$ErrorActionPreference = "SilentlyContinue" # Toggle Error Output. "SilentlyContinue" suppresses ERRORS, Continue shows all errors in job data

# FUNCTIONS START
Function Test-PathEx
{
    Param($Path)

    If (Test-Path $Path)
    {
        $True
    }
    Else
    {
        $Parent = Split-Path $Path
        [System.IO.Directory]::EnumerateFiles($Parent) -Contains $Path
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

    }
    Process
    {
        $FreeSpace = Get-WmiObject Win32_LogicalDisk -ComputerName $ComputerName |
        Where-Object { $_.DeviceID -eq "$DriveLetter" } |
        Select-Object @{Name="ComputerName"; Expression={ $_.SystemName } }, @{Name="DriveLetter"; Expression={ $_.Caption } }, @{Name="FreeSpace"; Expression={ "$([math]::Round($_.FreeSpace / 1GB,2))GB" } }, @{Name="PercentFree"; Expression={"$([math]::Round($_.FreeSpace / $_.Size,2) * 100)%"}} | Format-List
    }
    End
    {
        Return $FreeSpace
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
        $CombinedPath = Join-Path -Path "\\$ComputerName" -ChildPath "$DriveLetter`$" | Join-Path -ChildPath "$Path"

        # Start progress bar
        Write-Progress -Id 1 -Activity "Enumerating files on $ComputerName" -PercentComplete 0
    }

    Process
    {
        # Progress Bar counter
        $CurrentFileCount = 0
        $CurrentFolderCount = 0

        # Enumerate files, silence errors
        $Files = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -Recurse -ErrorAction SilentlyContinue -Attributes !Directory) | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
        # Timer Stop

        # Total file count for progress bar
        $FileCount = ($Files | Measure-Object).Count
        $TotalSize = ($Files | Measure-Object -Sum Length).Sum
        $TotalSize = [math]::Round($TotalSize / 1GB,3)

        "Removing $FileCount files... $TotalSize`GB."

        If ($DeleteFailed)
        {
            Remove-Variable DeleteFailed
        }
        $DeleteFailed = @()
        $DeleteFail = $False
        ForEach ($File in $Files)
        {
            $CurrentFileCount++
            $FullFileName = $File.FullName
            $Percentage = [math]::Round(($CurrentFileCount / $FileCount) * 100)
            Write-Progress -Id 1 -Activity "Removing Files" -CurrentOperation "File: $FullFileName" -PercentComplete $Percentage -Status "Progress: $CurrentFileCount of $FileCount, $Percentage%"
            Write-Verbose "Removing file $FullFileName"
            Try
            {
                $File | Remove-Item -Force
            }
            Catch
            {
                $DeleteFail = $True
                $DeleteFailed += "$FullFileName delete failed"
            }
        }

        # Reset progress bar for phase 2
        Write-Progress -Id 1 -Activity "Enumerating empty directories on $ComputerName" -CurrentOperation "Path: $CombinedPath" -PercentComplete 0

        # Enumerate remaining files
        $RemainingFiles = @(Get-ChildItem -Force -Path "$CombinedPath" -Recurse -ErrorAction SilentlyContinue -Attributes !Directory).Count
        # Enumerate folders with 0 files
        $EmptyFolders = @(Get-ChildItem -Force -Path "$CombinedPath" -Recurse -Attributes Directory) | Where-Object {($_.GetFiles()).Count -eq 0} | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
    
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
                    Write-Verbose "$FullFolderName not empty, skipping..."
                    Continue
                }
                Write-Progress -Id 1 -Activity "Removing Empty Directories" -CurrentOperation "Removing Empty Directory: $FullFolderName" -PercentComplete "$Percentage" -Status "Progress: $CurrentFolderCount of $EmptyCount, $Percentage%"
                Write-Verbose "Removing folder $FullFolderName"
                Try
                {
                    $EmptyFolder | Remove-Item -Force
                }
                Catch
                {
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
If (!(Get-Module ActiveDirectory))
{
    Import-Module -Name ActiveDirectory -ErrorAction Stop
}

# Prepare logging
$LogDate = (Get-Date).ToString('yyyy-MM-dd')
# Make log directory if it doesn't exist
If (!(Test-PathEx "$PSScriptRoot\logs\servercleanup"))
{
    New-Item -ItemType Directory "$PSScriptRoot\logs\servercleanup"
}
$LogPath = "$PSScriptRoot\logs\servercleanup"

# Prepare array and import CSV
$ServerList = @()
Write-Host "Importing $ServerListCSV"
Try
{
    Import-CSV -LiteralPath "$PSScriptRoot\$ServerListCSV" | ForEach-Object {$ServerList += $_."$CSVHeader"}
}
Catch
{
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
    If ($UserProfiles)
    {
        Remove-Variable UserProfiles
    }
    If ($DeleteMethodList)
    {
        Remove-Variable DeleteMethodList
    }
    If ($FailureList)
    {
        Remove-Variable FailureList
    }
 
   # Create UserProfiles, DeleteMethodList,and FailureList arrays
    $UserProfiles = @()
    $DeleteMethodList = @()
    $FailureList = @()

    # Progress bar
    $Counter++
    $Percentage = [math]::Round(($Counter / $TotalServers) * 100)
    Write-Progress -Activity "Recovering Disk Space on Servers" -CurrentOperation "Server Name: $Server" -Id 0 -PercentComplete $Percentage -Status "$Counter of $TotalServers"

    If (Test-Connection $Server)
    {
        $ServerName = Get-WmiObject Win32_ComputerSystem -ComputerName $Server
        $ServerHostName = $ServerName.__SERVER
    }
    Else
    {
        Write-Warning "$Server did not respond, skipping" | Tee-Object -FilePath "$LogPath\skipped-$LogDate.log" -Append
        Continue
    }

    # Free space
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter "C:" | Tee-Object -FilePath "$LogPath\$ServerHostName-freespace-$LogDate.log" -Append
    #Start-Sleep -Seconds 2

    # Get profile list excluding service accounts and the local administrator account
    $UserProfiles += Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "NOT Special='True' AND NOT LocalPath LIKE '%00%' AND NOT LocalPath LIKE '%Administrator' AND NOT LocalPath LIKE '%SQL%' AND NOT LocalPath LIKE '%MSSQL%' AND NOT LocalPath LIKE '%Classic .NET AppPool' AND NOT LocalPath LIKE '%.NET%' AND NOT LocalPath LIKE '%MsDts%' AND NOT LocalPath LIKE '%ReportServer%'"

    ForEach ($UserProfile in $UserProfiles)
    {
        $UserLocalPath = $UserProfile.LocalPath.Split("\")[-1]
        If ($SID)
        {
            Remove-Variable SID
        }
        # If profile status Bit Field includes 8 (corrupt profile), quit.
        If ((8 -band $UserProfile.Status) -eq 8)
        {
            Write-Host "Checking for local profile corruption..."
            Write-Warning "PROFILE CORRUPT!"
            Write-Host "Flagged `"$UserLocalPath`" for removal." -ForegroundColor Yellow
            $DeleteMethodList += $UserProfile
            Continue
        }
        Else
        {
            "Profile `"{0}`"" -F $UserLocalPath
        }
        $SID = $UserProfile | Select-Object -ExpandProperty sid
        Try
        {
            Get-ADUser -Identity $SID | Out-Null
        }
        Catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
        {
            Write-Warning "Profile $UserLocalPath {SID: $SID} does not exist in Active Directory."
            Write-Host "Flagged `"$UserLocalPath`" for removal."
            $DeleteMethodList += $UserProfile
            Continue
        }
        Catch
        {
            "Unhandled Exception with Active Directory module, skipping $UserLocalPath."
            Continue
        }
        Write-Host "$UserLocalPath - Status OK" -ForegroundColor Green
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
            $FailureList += $User.LocalPath
            $FailureCount++
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
        Write-Host "$FailureCount profile(s) failed to delete. Check `"$LogPath\$ServerHostName-userprofiles-$LogDate.log`" for details."
        $FailureList | Out-File -LiteralPath "$LogPath\$ServerHostName-userprofiles-$LogDate.log"
    }

    # Set drive letter for cleanup
    $DriveLetter = "C"

    # C:\$RECYCLE.BIN
    $RelativePath = '$RECYCLE.BIN'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
    If (Test-PathEx "$PathTest")
    {
        Write-Host "Emptying $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\MSOCache
    $RelativePath = 'MSOCache'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
    If (Test-PathEx "$PathTest")
    {
        Write-Host "Emptying $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\Temp
    $RelativePath = 'Windows\Temp'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
    If (Test-PathEx "$PathTest")
    {
        Write-Host "Emptying $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Temp
    $RelativePath = 'Temp'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
    If (Test-PathEx "$PathTest")
    {
        Write-Host "Emptying $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\ProPatches\Patches
    $RelativePath = 'Windows\ProPatches\Patches'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
    If (Test-PathEx "$PathTest")
    {
        Write-Host "Emptying $PathTest"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    $RemainingProfiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "NOT Special='True' AND NOT LocalPath LIKE '%00%' AND NOT LocalPath LIKE '%Administrator' AND NOT LocalPath LIKE '%SQL%' AND NOT LocalPath LIKE '%MSSQL%' AND NOT LocalPath LIKE '%Classic .NET AppPool' AND NOT LocalPath LIKE '%Default%'"

    ForEach ($UserProfile in $RemainingProfiles)
    {
        $UserPath = $UserProfile.LocalPath
        $UserPath = $UserPath -replace '[C|c]:\\',''

        # User temp files
        $RelativePath = "$UserPath\AppData\Local\Temp"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User IE Cache (New)
        $RelativePath = "$UserPath\AppData\Local\Microsoft\Windows\INetCache"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User IE Cache (Old)
        $RelativePath = "$UserPath\AppData\Local\Microsoft\Windows\Temporary Internet Files"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Chrome Cache
        $RelativePath = "$UserPath\AppData\Local\Google\Chrome\User Data\Default\Cache"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Chrome Updates
        $RelativePath = "$UserPath\AppData\Local\Google\Chrome\Update"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
        # User Crash Dumps
        $RelativePath = "$UserPath\AppData\Local\CrashDumps"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath\"
        If (Test-PathEx "$PathTest")
        {
            Write-Host "Emptying $PathTest"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
    }
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter "C:" | Tee-Object -FilePath "$LogPath\$ServerHostName-freespace-$LogDate.log" -Append
    #Start-Sleep -Seconds 1
    Write-Host "Cleanup completed on $ServerHostName"
}
Write-Progress -Id 0 "Done" "Done"
Write-Host "Job complete. Check $LogPath for log files"

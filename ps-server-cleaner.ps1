#requires -Version 3.0
#requires -RunAsAdministrator
# ps-server-cleaner.ps1
# Cleans servers before and after patch night
# Contact ExEvolution http://www.reddit.com/user/exevolution if you have any issues
Clear-Host

# Configuration options
$ServerListCSV = "serverlist.csv" # Filename of CSV listing all servers to be cleaned, should be placed in the same directory as this script
$CSVHeader = "Hostname" # Header of the column in the CSV file containing the machines DNS name or IP address
$VerbosePreference = "SilentlyContinue" # Toggle Verbosity, "SilentlyContinue" to suppress VERBOSE messages, "Continue" to use full Verbosity
$ErrorActionPreference = "SilentlyContinue" # Toggle Error Output. "SilentlyContinue" suppresses ERRORS, Continue shows all errors in job data

# FUNCTIONS START
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
        Select-Object @{Name="ComputerName"; Expression={ $_.SystemName } }, @{Name="DriveLetter"; Expression={ $_.Caption } }, @{Name="FreeSpace"; Expression={ "$([math]::Round($_.FreeSpace / 1GB,2))GB" } }, @{Name="PercentFree"; Expression={"$([math]::Round($_.FreeSpace / $_.Size,2) * 100)%"}}
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
    }

    Process
    {
        # Progress Bar counter
        $CurrentFileCount = 0
        $CurrentFolderCount = 0
    
        ''
        '--------------------------------------------------'
        "Enumerating files on $ComputerName, please wait..."
    
        # Start progress bar
        Write-Progress -Id 0 -Activity "Enumerating files to remove $ComputerName" -PercentComplete 0

        # Timer Start
        $T0 = Get-Date

        # Enumerate files, silence errors
        $Files = @(Get-ChildItem -Force -LiteralPath "$CombinedPath" -Recurse -ErrorAction SilentlyContinue -Attributes !Directory) | Sort-Object -Property @{ Expression = {$_.FullName.Split('\').Count} } -Descending
        # Timer Stop
        $T1 = Get-Date
        $T2 = New-TimeSpan -Start $T0 -End $T1
        "Operation Completed in {0:d2}:{1:d2}:{2:d2}" -F $T2.Hours,$T2.Minutes,$T2.Seconds

        # Total file count for progress bar
        $FileCount = ($Files | Measure-Object).Count
        $TotalSize = ($Files | Measure-Object -Sum Length).Sum
        $TotalSize = [math]::Round($TotalSize / 1GB,3)

        "Removing $FileCount files... $TotalSize`GB."

        # Timer Start
        $T0 = Get-Date

        If ($DeleteFailed)
        {
            Remove-Variable DeleteFailed
        }
        $DeleteFailed = @()
        ForEach ($File in $Files)
        {
            $CurrentFileCount++
            $FullFileName = $File.FullName
            $Percentage = [math]::Round(($CurrentFileCount / $FileCount) * 100)
            Write-Progress -Id 0 -Activity "Removing $Title" -CurrentOperation "File: $FullFileName" -PercentComplete $Percentage -Status "Progress: $CurrentFileCount of $FileCount, $Percentage%"
            Write-Verbose "Removing file $FullFileName"
            Try
            {
                $File | Remove-Item -Force
            }
            Catch
            {
                $DeleteFailed += "$FullFileName delete failed"
            }
        }
        Write-Progress -Id 0 -Completed -Activity 'Done'

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
                    $DeleteFailed += "$FullFolderName folder delete failed"
                }
            }
            Write-Progress -Id 1 -Completed -Activity 'Done'
        }
    }

    End
    {
        $DeleteFailed | Out-File -LiteralPath "$LogPath\$ServerHostName-files-$LogDate.log" -Append
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
If (!(Test-Path -LiteralPath "$PSScriptRoot\logs\servercleanup"))
{
    New-Item -ItemType Directory "$PSScriptRoot\logs\servercleanup"
}
$LogPath = "$PSScriptRoot\logs\servercleanup"

# Prepare array and import CSV
$ServerList = @()
Write-Host "Importing $ServerListCSV"
Import-CSV -LiteralPath "$PSScriptRoot\$ServerListCSV" | ForEach-Object {$ServerList += $_."$CSVHeader"}
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
    Start-Sleep -Seconds 2

    # Get profile list excluding service accounts and the local administrator account
    $UserProfiles += Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "NOT Special='True' AND NOT LocalPath LIKE '%00%' AND NOT LocalPath LIKE '%Administrator' AND NOT LocalPath LIKE '%SQL%' AND NOT LocalPath LIKE '%MSSQL%' AND NOT LocalPath LIKE '%Classic .NET AppPool'"

    ForEach ($UserProfile in $UserProfiles)
    {
        If ($SID)
        {
            Remove-Variable SID
        }
        # If profile status Bit Field includes 8 (corrupt profile), quit.
        If ((8 -band $UserProfile.Status) -eq 8)
        {
            Write-Host "Checking for local profile corruption..."
            Write-Warning "PROFILE CORRUPT!"
            "Flagged `"{0}`" for removal." -F $UserProfile.LocalPath.Split("\")[-1]
            $DeleteMethodList += $UserProfile
            Continue
        }
        Else
        {
            "Profile `"{0}`" - Status OK" -F $UserProfile.LocalPath.Split("\")[-1]
        }
        $SID = $UserProfile | Select-Object -ExpandProperty sid
        Try
        {
            Get-ADUser -Identity $SID | Out-Null
        }
        Catch
        {
            Write-Warning "Profile does not exist in Active Directory"
            Write-Warning "$SID"
            "Flagged `"{0}`" for removal with DelProf2." -F $UserProfile.LocalPath.Split("\")[-1]
            $DeleteMethodList += $UserProfile
            Continue
        }
    }
    "Scanned all profiles on $ServerHostName"
    "{0} of {1} local profiles scheduled for deletion on {2}" -F ($DeleteMethodList | Measure-Object).Count,($UserProfiles | Measure-Object).Count,$ServerHostName

    ForEach ($User in $DeleteMethodList)
    {
        "Deleting Profile: {0}" -F $User.LocalPath.Split("\")[-1]
        Try
        {
            $User.Delete()
            "Success!"
        }
        Catch
        {
            "Deletion failed."
            $FailureList += $User.LocalPath.Split("\")[-1]
        }
    }
    "Unused profiles removed"
    $FailureList | Out-File -LiteralPath "$LogPath\$ServerHostName-userprofiles-$LogDate.log"

    # Set drive letter for cleanup
    $DriveLetter = "C"

    # C:\$RECYCLE.BIN
    Write-Host "Checking C:\Recycle Bin"
    $RelativePath = '$RECYCLE.BIN'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If (Test-Path "$PathTest")
    {
        Write-Host "Emptying C:\Recycle Bin"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\MSOCache
    Write-Host "Checking C:\MSOCache"
    $RelativePath = 'MSOCache'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If (Test-Path "$PathTest")
    {
        Write-Host "Emptying C:\MSOCache"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\Temp
    Write-Host "Checking C:\Windows\Temp"
    $RelativePath = 'Windows\Temp'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If (Test-Path "$PathTest")
    {
        Write-Host "Emptying C:\Windows\Temp"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    # C:\Windows\ProPatches\Patches
    Write-Host "Checking C:\Windows\ProPatches\Patches"
    $RelativePath = 'Windows\ProPatches\Patches'
    $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
    If (Test-Path "$PathTest")
    {
        Write-Host "Emptying C:\Windows\ProPatches\Patches"
        Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
    }

    $RemainingProfiles = Get-WmiObject -Class Win32_UserProfile -ComputerName $ServerHostName -Filter "NOT Special='True' AND NOT LocalPath LIKE '%00%' AND NOT LocalPath LIKE '%Administrator' AND NOT LocalPath LIKE '%SQL%' AND NOT LocalPath LIKE '%MSSQL%' AND NOT LocalPath LIKE '%Classic .NET AppPool'"

    ForEach ($UserProfile in $RemainingProfiles)
    {
        $UserPath = $UserProfile.LocalPath
        $UserPath = $UserPath -replace '[C|c]:\\',''
        # User temp files
        Write-Host "Checking C:\$UserPath"
        $RelativePath = "$UserPath\AppData\Local\Temp"
        $PathTest = "\\$ServerHostName\$DriveLetter`$\$RelativePath"
        If (Test-Path "$PathTest")
        {
            Write-Host "Emptying C:\$UserPath\AppData\Local\Temp"
            Remove-WithProgress -ComputerName $ServerHostName -DriveLetter $DriveLetter -Path $RelativePath
        }
    }
    Get-FreeSpace -ComputerName $ServerHostName -DriveLetter "C:" | Tee-Object -FilePath "$LogPath\$ServerHostName-freespace-$LogDate.log" -Append
    Start-Sleep -Seconds 1
    "Cleanup completed on $ServerHostName, moving to next system"
}
Write-Progress -Id 0 "Done" "Done"

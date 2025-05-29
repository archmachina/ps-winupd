<#
#>

########
# Global settings
$ErrorActionPreference = "Stop"
$InformationPreference = "Continue"
Set-StrictMode -Version 2

Function Update-WinUpdCabFile
{
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Uri = "https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab",

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [switch]$Force = $false
    )

    process
    {
        # Modification times for comparison
        $cabModificationTime = $null
        $urlModificationTime = $null

        # Determine modification time for the local cab file
        try {
            Write-Verbose "Getting current cab file modification time"
            $cabModificationTime = (Get-Item $Path).LastWriteTimeUtc

            Write-Verbose "Cab Modification Time: $cabModificationTime"
        } catch {
            Write-Verbose "Could not get current cab file modification time: $_"
        }

        # Determine modification time for the remote cab file
        Write-Verbose "Retrieving URL modification time"
        $params = @{
            UseBasicParsing = $true
            Method = "Head"
            Uri = $Uri
        }

        $headers = Invoke-WebRequest @params

        # Extract modification time
        try {
            $urlModificationTime = [DateTime]::Parse($headers.Headers["Last-Modified"])

            Write-Verbose "Url Modification Time: $urlModificationTime"
        } catch {
            Write-Warning "Failed to parse Last-Modified header as DateTime: $_"
        }

        # If we have cab and url modification times and the local file is newer than the url, then finish here
        if (!$Force -and $null -ne $cabModificationTime -and $null -ne $urlModificationTime -and $cabModificationTime -gt $urlModificationTime)
        {
            Write-Verbose "Local file is newer than the Uri. Not downloading."
            return
        }

        Write-Verbose "wsusscn2.cab file needs updating. Downloading."

        # Temporary path for download file
        $tempPath = $Path + ".tmp"
        if (Test-Path $tempPath)
        {
            Remove-Item -Force $tempPath
        }

        # Download to temporary location
        $params = @{
            UseBasicParsing = $true
            Method = "Get"
            Uri = $Uri
            OutFile = $tempPath
        }

        Invoke-WebRequest @params

        # Move the temporary location to the actual place for the scn2 cab file
        Move-Item $tempPath $Path -Force

        Write-Verbose "Successfully updated wsusscn2.cab"
    }
}

Function Remove-WinUpdOfflineScan
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$Name = "Offline Sync Service"
    )

    Process
    {
        # Create the ServiceManager object
        Write-Verbose "Creating ServiceManager object"
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        # Remove any services by this name
        $services = $manager.Services | Where-Object { $_.Name -eq $Name } | ForEach-Object { $_ }
        $services | ForEach-Object {
            $service = $_

            Write-Verbose ("Removing service with ID: " + $service.ServiceID)
            $manager.RemoveService($service.ServiceID)
        }
    }
}

Function Update-WinUpdOfflineScan
{
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$OfflineServiceName = "Offline Sync Service",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$CabFile = $null,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [switch]$Force = $false
    )

    process
    {
        # Create the service manager to interact with Windows Update
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        # Collect all package services matching the name
        [array]$services = @($manager.Services | ForEach-Object { $_ })
        $services = $services |
            Where-Object { $_.Name -eq $OfflineServiceName } |
            Sort-Object -Property IssueDate -Descending
        Write-Verbose "Current package services:"
        Write-Verbose ($services | Format-Table -Property Name,ServiceId,IssueDate | Out-String)

        # Remove all but the newest
        $services | Select-Object -Skip 1 | ForEach-Object {
            $service = $_

            Write-Verbose ("Removing service with ID: " + $service.ServiceID)
            $manager.RemoveService($service.ServiceID)
        }

        # Get details for the active service, if there is one
        $currentIssueDate = $null
        $currentServiceId = $null
        if (($services | Measure-Object).Count -gt 0)
        {
            #$currentIssueDate = [DateTime]::SpecifyKind($services[0].IssueDate, [DateTimeKind]::Utc)
            $currentIssueDate = $services[0].IssueDate
            $currentServiceId = $services[0].ServiceID

            Write-Verbose ("Current service issue date: " + $currentIssueDate.ToString("o"))
            Write-Verbose ("Current service ID: " + $currentServiceId)
        }

        # Get the full path to the cab file
        $CabFile = (Get-Item $CabFile).FullName
        Write-Verbose "Cab File path: $CabFile"

        # Get the cab file modification time
        # Note - IssueDate is in UTC
        $cabModificationTime = (Get-Item $CabFile).LastWriteTimeUtc
        Write-Verbose ("Cab modification time: " + $cabModificationTime.ToString("o"))

        # Add the offline scan file to Windows Update
        if ($Force -or $null -eq $currentIssueDate -or $cabModificationTime -gt $currentIssueDate)
        {
            Write-Verbose "Updating searcher to use the local cab file"

            # Add the cab file to Windows Update
            $service = $manager.AddScanPackageService($OfflineServiceName, $CabFile, 1)

            if ($null -ne $currentServiceId)
            {
                # Remove the now defunct service
                Write-Verbose "Removing defunct service: $currentServiceId"
                $manager.RemoveService($currentServiceId)
            }

            # Return the new service ID
            Write-Verbose ("New service ID: {0}" -f $service.ServiceId)
            $service.ServiceId
        } else {
            Write-Verbose "Service does not need updating. Returning current service: $currentServiceId"

            # Return the current service ID
            $currentServiceId
        }
    }
}

Function Get-WinUpdScanServices
{
    [CmdletBinding()]
    param()

    Process
    {
        # Create the ServiceManager object
        Write-Verbose "Creating ServiceManager object"
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        $manager.Services
    }
}

Function Get-WinUpdUpdates
{
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$ServiceId = $null
    )

    process
    {
        # Create the ServiceManager object
        Write-Verbose "Creating ServiceManager object"
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        # Create an update searcher
        Write-Verbose "Creating update session using new scan service"
        $session = New-Object -ComObject Microsoft.Update.session
        $searcher = $session.CreateUpdateSearcher()

        # Use a scan service, if specified
        if (![string]::IsNullOrEmpty($ServiceId))
        {
            # Update the searcher to use the cab file
            Write-Verbose "Updating searcher to use the local cab file"
            $searcher.ServerSelection = 3
            $searcher.ServiceID = $ServiceId
        }

        # Capture anything not installed
        Write-Verbose "Searching for packages that are not installed"
        $result = $searcher.Search("IsInstalled=0")
        $count = ($result.Updates | Measure-Object).Count
        Write-Verbose "Found $count updates not installed"

        # Report on any updates found
        $result.Updates
    }
}

Function Install-WinUpdUpdates
{
    param(
        [Parameter(Mandatory=$true)]
        $Updates,

        [Parameter(Mandatory=$false)]
        [ValidateNotNull()]
        [switch]$DownloadOnly = $false
    )

    process
    {
        # Create a new Update Collection
        $updateCol = New-Object -ComObject Microsoft.Update.UpdateColl
        $Updates | ForEach-Object { $updateCol.Add($_) | Out-Null }

        # Report on applicable updates
        $count = ($updateCol | Measure-Object).Count
        Write-Verbose "$count updates to apply"

        # Quit here if there are no updates to apply
        if (($updateCol | Measure-Object).Count -eq 0)
        {
            return
        }

        # Download any packages
        Write-Verbose "Starting download of updates"
        $downloader = New-Object -ComObject Microsoft.Update.Downloader
        $downloader.Updates = $updateCol
        $downloader.Download() | Out-Null

        if (!$DownloadOnly)
        {
            # Install any packages
            Write-Verbose "Starting installation of updates"
            $installer = New-Object -ComObject Microsoft.Update.Installer
            $installer.ForceQuiet = $true
            $installer.Updates = $updateCol
            $installer.Install() | Out-Null

            Write-Verbose "Update installation completed"
        }
    }
}

Function Get-WinUpdRebootRequired
{
    param()

    process
    {
        $systemInfo = New-Object -ComObject Microsoft.Update.SystemInfo
        [bool]($systemInfo.RebootRequired)
    }
}


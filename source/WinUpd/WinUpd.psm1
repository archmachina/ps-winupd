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
        [string]$Uri = "https://catalog.s.download.windowsupdate.com/microsoftupdate/v6/wsusscan/wsusscn2.cab"
    )

    process
    {
        # Determine modification time for the local cab file
        $cabModificationTime = $null
        try {
            Write-Verbose "Getting current cab file modification time"
            $cabModificationTime = (Get-Item $Path).LastWriteTimeUtc
        } catch {
            Write-Verbose "Could not get current cab file modification time: $_"
        }

        Write-Verbose "Cab Modification Time: $cabModificationTime"

        # Determine modification time for the remote cab file
        Write-Verbose "Retrieving URL modification time"
        $params = @{
            UseBasicParsing = $true
            Method = "Head"
            Uri = $Uri
        }

        $headers = Invoke-WebRequest @params

        # Extract modification time
        $urlModificationTime = $null
        try {
            $urlModificationTime = [DateTime]::Parse($headers.Headers["Last-Modified"])
        } catch {
            Write-Warning "Failed to parse Last-Modified header as DateTime: $_"
        }

        Write-Verbose "Url Modification Time: $urlModificationTime"

        # Download the file, if it's newer than what we have locally or we don't have valid modification
        # time information
        if ($null -eq $cabModificationTime -or $null -eq $urlModificationTime -or $urlModificationTime -gt $cabModificationTime)
        {
            Write-Verbose "wsusscn2.cab file needs updating. Downloading."
            $params = @{
                UseBasicParsing = $true
                Method = "Get"
                Uri = $Uri
                OutFile = ("{0}.tmp" -f $Path)
            }

            # Download to temporary location
            Remove-Item -Force ("{0}.tmp" -f $Path) -EA Ignore
            Invoke-WebRequest @params

            # Move the temporary location to the actual place for the scn2 cab file
            Move-Item ("{0}.tmp" -f $Path) $Path -Force

            Write-Verbose "Successfully downloaded wsusscn2.cab"
        } else {
            Write-Verbose "Local file is newer than the Uri. Not downloading."
        }
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

Function New-WinUpdOfflineScan
{
    param(
        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$OfflineServiceName = "Offline Sync Service",

        [Parameter(Mandatory=$false)]
        [ValidateNotNullOrEmpty()]
        [string]$CabFile = $null
    )

    process
    {
        # Get the full path to the cab file
        $CabFile = (Get-Item $CabFile).FullName
        Write-Verbose "Cab File path: $CabFile"

        # Remove any existing instances of the package service with this name
        Write-Verbose "Removing any preexisting service registrations for `"$OfflineServiceName`""
        Remove-WinUpdOfflineScan -Name $OfflineServiceName

        # Create the ServiceManager object
        Write-Verbose "Creating ServiceManager object"
        $manager = New-Object -ComObject Microsoft.Update.ServiceManager

        # Update the searcher to use the cab file
        Write-Verbose "Updating searcher to use the local cab file"
        $service = $manager.AddScanPackageService($OfflineServiceName, $CabFile, 1)

        # Return the service ID
        $service.ServiceId
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

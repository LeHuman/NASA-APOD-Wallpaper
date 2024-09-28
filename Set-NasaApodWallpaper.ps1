<#
.SYNOPSIS
    NASA APOD Auto Wallpaper script
.DESCRIPTION
    PowerShell Script to automatically download the latest NASA APOD image and set it as the windows background image. Run with no parameters to setup.
.LINK
    https://github.com/lehuman/NASA-APOD-Wallpaper
#>

param(
    # Only update the wallpaper. Don't ask for input unless needed. Will not setup a task.
    [switch]$Update,
    # Do not add any text to the wallpaper.
    [switch]$NoText,
    # Do not crop the wallpaper to monitor ratio, text may not be visible.
    [switch]$NoCrop,
    # Do not move the script to the APOD folder. Make sure the script is not moved after setup.
    [switch]$NoMove,
    # Do not install anything requring admin rights, reducing functionality. UAC prompt will still occur for task setup.
    [switch]$NoAdmin,
    # Do not download or install anything extra. Will set other flags as needed, reducing functionality.
    # A console will appear when scheduled task is run.
    [switch]$NoDownload,
    # Run silently, will not wait for user input or output anything. Will attempt to install anything necessary.
    # May still trigger UAC prompt for setup.
    # Consoles will appear on first set-up to show installation of dependencies.
    [switch]$Silent,
    # Force actions where possible, such as redownloading the APOD
    [switch]$Force
)

####### OPTIONS #########

# Arguments to pass for silent mode
if ($Silent) {
    $SilentArgs = "-WindowStyle Hidden -NonInteractive"
}
else {
    $SilentArgs = ""
}
# Argument list to pass to admin setup
$optionsPassed = $MyInvocation.Line -replace "^.*$($MyInvocation.MyCommand.Name)\s*".Trim()
# Arguments to pass to admin setup
$arguments = "$SilentArgs -ExecutionPolicy Bypass & '" + $MyInvocation.MyCommand.Definition + "' $optionsPassed"
# Images will be downloaded and saved to this folder
$downloadDir = "$([Environment]::GetFolderPath("MyPictures"))\APOD"
# The path of the script, will ask to move if != $downloadDir
$scriptDir = $PSScriptRoot
# The name of the script
$scriptName = $MyInvocation.MyCommand.Name
# Name of scheduled task
$taskName = "NASA APOD BG Image Updater"
# Options for RunHidden.exe
$RunHidden = $true
$RunHiddenName = "RunHidden.exe"
$RunHiddenURL = "https://github.com/LesFerch/RunHidden/releases/latest/download/RunHidden.zip"

if ($NoDownload) {
    $NoCrop = $true
    $NoText = $true
    $RunHidden = $false
}

### LIBRARY FUNCTIONS ###

# Get Monitor Resolution
# https://superuser.com/questions/1436909/how-to-get-the-current-screen-resolution-on-windows-via-command-line
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class PInvoke {
    [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
}
"@

function Get-Resolution() {
    $hdc = [PInvoke]::GetDC([IntPtr]::Zero)
    $width = [PInvoke]::GetDeviceCaps($hdc, 118)
    $height = [PInvoke]::GetDeviceCaps($hdc, 117)
    return $width, $height
}

# Manually Set Wallpaper
$setwallpapersrc = @"
using System.Runtime.InteropServices;

public class Wallpaper {
    public const int SetDesktopWallpaper = 20;
    public const int UpdateIniFile = 0x01;
    public const int SendWinIniChange = 0x02;
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
    public static void SetWallpaper(string path) {
        SystemParametersInfo(SetDesktopWallpaper, 0, path, UpdateIniFile | SendWinIniChange);
    }
}
"@
Add-Type -TypeDefinition $setwallpapersrc

function Set-Wallpaper($imagePath) {
    [Wallpaper]::SetWallpaper($imagePath)
}

# Wrapper function for Write-Host
function Write-SafeHost {
    param (
        # [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter()]
        [ConsoleColor]$ForegroundColor,
        [ConsoleColor]$BackgroundColor
    )

    if (-not $Silent) {
        # If Silent is not enabled, write the message
        if ($ForegroundColor -and $BackgroundColor) {
            Write-Host -ForegroundColor $ForegroundColor -BackgroundColor $BackgroundColor $Message
        }
        elseif ($ForegroundColor) {
            Write-Host -ForegroundColor $ForegroundColor $Message
        }
        else {
            Write-Host $Message
        }
    }
}

####### FUNCTIONS #######

# Check if the script is already in the target directory
function Get-ScriptInDir {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory
    )
    return $scriptDir -eq $TargetDirectory
}

# Run script from APOD folder, returning if already so or if failed to do so.
function Move-ScriptAndRelaunch {
    param (
        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory
    )

    # Get the full path of the current script
    $targetScriptPath = Join-Path -Path $scriptDir -ChildPath $scriptName
    $targetScriptDest = Join-Path -Path $TargetDirectory -ChildPath $scriptName

    # Make sure the target directory exists
    if (-not (Test-Path -Path $TargetDirectory)) {
        Write-SafeHost "Creating target directory: $TargetDirectory"
        New-Item -ItemType Directory -Path $TargetDirectory -Force
    }

    # Check if the script is already in the target directory
    if (Get-ScriptInDir -TargetDirectory $TargetDirectory) {
        Write-SafeHost "Script is already running from the target directory: $TargetDirectory"
        return $true
    }

    # Move the script to the target directory
    try {
        Write-SafeHost "Moving script to $TargetDirectory..."
        Move-Item -Path $targetScriptPath -Destination $targetScriptDest -Force
        Write-SafeHost "Script moved successfully."
    }
    catch {
        Write-SafeHost "Error moving script: $_"
        return $false
    }

    # Relaunch the script from the new location
    try {
        Write-SafeHost "Relaunching script from $TargetDirectory..."
        Start-Process -FilePath "powershell.exe" -ArgumentList "$SilentArgs -ExecutionPolicy Bypass -File `"$targetScriptDest`" $optionsPassed"
        Write-SafeHost "Script relaunched successfully."
    }
    catch {
        Write-SafeHost "Error relaunching script: $_"
        return $false
    }

    # Exit the original process after relaunching
    Write-SafeHost "Exiting"
    exit
}

# Check if the script is running with admin rights, relaunching if not
function Get-AdminRights {
    if ($NoAdmin) {
        return $false
    }
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {  
        try {
            Start-Process powershell -Verb runAs -ArgumentList $arguments
            exit
        }
        catch {
            Write-Warning "Failed to elevate the script to administrator privileges. The script requires admin rights to run. Use -NoAdmin to disallow installation of admin dependencies. Exiting."
            exit 1
        }
    }
    return $true
}

# Install system dependencies
function Install-Dependencies {
    # First Check for PowerBGInfo
    $module = Get-InstalledModule -Name "PowerBGInfo" -ErrorAction SilentlyContinue

    # If PowerBGInfo not here, get package manager
    if (-not $module) {
        # Check and install NuGet if not present
        $nugetProvider = Get-PackageProvider -Name "NuGet" -ErrorAction SilentlyContinue
        if (-not $nugetProvider -or ($nugetProvider.Version -lt [version]"2.8.5.201")) {
            if (-not (Get-AdminRights)) {
                return $false
            }
            Write-Host "APOD is Installing the NuGet Package Provider to install dependencies"
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Verbose
            Write-Host "Done"
            Start-Sleep 1
        }
    }

    # Install PowerBGInfo module if not present
    if (-not $module) {
        if (-not (Get-AdminRights)) {
            return $false
        }
        Write-Host "APOD is Installing the PowerBGInfo Module to crop and apply text to the wallpaper"
        Install-Module -Name PowerBGInfo -Force -Verbose
        Write-Host "Done"
        Start-Sleep 1
    }

    return $true
}

# Download and unzip a link to path
function Get-UnzipFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$DownloadUrl,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    # Ensure destination path exists
    if (-not (Test-Path -Path $DestinationPath)) {
        New-Item -ItemType Directory -Path $DestinationPath -Force
    }

    # Define the zip file path
    $zipFilePath = Join-Path -Path $DestinationPath -ChildPath "$FileName.zip"
    $unzippedFilePath = Join-Path -Path $DestinationPath -ChildPath $FileName

    # Check if the unzipped folder already exists
    if (Test-Path -Path $unzippedFilePath) {
        Write-SafeHost "The file '$unzippedFilePath' already exists."
        return $true
    }

    # Download the file
    try {
        Write-SafeHost "Downloading file..."
        $client = New-Object System.Net.WebClient
        $client.DownloadFile($DownloadUrl, $zipFilePath)
        Write-SafeHost "Download completed."
    }
    catch {
        Write-SafeHost "Error downloading file: $_"
        return $false
    }

    # Unzip the file
    try {
        Write-SafeHost "Unzipping file..."
        Expand-Archive -LiteralPath $zipFilePath -DestinationPath $DestinationPath -Force
        Write-SafeHost "Unzip completed."
    }
    catch {
        Write-SafeHost "Error unzipping file: $_"
        return $false
    }

    # Delete the zip file
    try {
        Remove-Item -LiteralPath $zipFilePath -Force
        Write-SafeHost "Deleted zip file."
    }
    catch {
        Write-SafeHost "Error deleting zip file: $_"
        return Test-Path -Path $unzippedFilePath
    }

    Write-SafeHost "Downloaded and unzipped successfully."
    return $true
}

# Check if the current network has internet access
function Test-NetworkConnection {
    $gateway = (Get-NetRoute -DestinationPrefix '0.0.0.0/0').NextHop
    if ($gateway) {
        return Test-Connection -ComputerName $gateway -Count 1 -ErrorAction SilentlyContinue
    }
    else {
        return $false
    }
}

# Define a function to check internet connectivity
function Test-InternetConnection {
    try {
        $request = [System.Net.WebRequest]::Create("ping.ubnt.com")
        $request.Timeout = 5000
        $request.GetResponse()
        return $true
    }
    catch {
        return $false
    }
}

# Wait for a network connection
function Wait-ForConnection {
    while (-not (Test-NetworkConnection) -and -not (Test-InternetConnection)) {
        Write-SafeHost "Waiting for internet..." -ForegroundColor Yellow
        Start-Sleep -Seconds 5
    }   
}

# Download current APOD image and title
function Get-CurrentApodImage() {
    # create download dir if it doesn't exist
    if (-not (Test-Path -LiteralPath $downloadDir)) {
        New-Item -ItemType Directory -Path $downloadDir
    }

    # Get current APOD page
    $url = [System.Uri]"https://apod.nasa.gov/apod/astropix.html"

    $segments = $url.Segments | Select-Object -SkipLast 1
    $base = "$($url.Scheme)://$($url.Host)$($segments -join `"`")"

    Wait-ForConnection
    $result = Invoke-WebRequest -Uri $url -TimeoutSec 10 -ErrorAction Stop

    $srcpattern = '(?i)href="(.*?)"'
    $titlepattern = '(?i)<b>(.*?)</b>'
    $src = ([regex]$srcpattern ).Matches($result.Content)
    $title_src = ([regex]$titlepattern ).Match($result.Content)

    $image = $null
    $title = $null

    # Get Title
    if (![System.String]::IsNullOrEmpty($title_src)) {
        $title = $title_src.Groups[1].Value.Trim()
    }

    # Get image url
    $src | ForEach-Object {
        $value = $_.Groups[1].Value.Trim().ToLower()
        if ($value.EndsWith(".jpg") -or $value.EndsWith(".jpeg") -or $value.EndsWith(".png")) {
            if ([System.String]::IsNullOrEmpty($image)) {
                $image = [System.Uri]($base + $_.Groups[1].Value)
            }
            else { return }
        }
    }

    if ($null -eq $image) {
        # no image found, abort
        return $null
    }

    $fileName = (Get-Date -Format "yyyyMMdd") + "_" + $image.Segments[$image.Segments.Count - 1]
    $fullDLPath = ($downloadDir + "\" + $fileName)

    # Check if file already exists
    if (Test-Path -LiteralPath $fullDLPath) {
        if ($Force) {
            Write-SafeHost -ForegroundColor Yellow "Force removing existing APOD"
            Remove-Item -Force $fullDLPath
        }
        else {
            Write-SafeHost -ForegroundColor Red "Current image already exists, aborting!"
            return $false;
        }
    }

    ### Download image
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $image -TimeoutSec 10 -ErrorAction Stop -OutFile $fullDLPath

    return $fullDLPath, $title
}

# Get a rectangle for cropping an image
function Get-CropRectangle($MonitorWidth, $MonitorHeight, $ImageWidth, $ImageHeight) {
    # Calculate aspect ratios
    $monitorAspectRatio = $MonitorWidth / $MonitorHeight
    $imageAspectRatio = $ImageWidth / $ImageHeight

    if ($imageAspectRatio -gt $monitorAspectRatio) {
        # Crop by width
        $cropHeight = $ImageHeight  # Full height
        $cropWidth = [math]::Floor($cropHeight * $monitorAspectRatio)  # Calculate new width
        # $cropX = [math]::Floor(($ImageWidth - $cropWidth) / 2)  # Center crop on x-axis
        # $cropY = 0  # Start from top
    }
    else {
        # Crop by height
        $cropWidth = $ImageWidth  # Full width
        $cropHeight = [math]::Floor($cropWidth / $monitorAspectRatio)  # Calculate new height
        # $cropX = 0  # Start from left
        # $cropY = [math]::Floor(($ImageHeight - $cropHeight) / 2)  # Center crop on y-axis
    }

    return $cropWidth, $cropHeight
}

# Function to run image update
function Update() {
    $imagePath, $title = Get-CurrentApodImage

    if ($imagePath -is [string] -and -not ([System.String]::IsNullOrEmpty($imagePath))) {
        Write-SafeHost "Setting wallpaper to $title"

        $hasDepend = Install-Dependencies

        if (-not $hasDepend) {
            $NoCrop = $true
            $NoText = $true
            Write-SafeHost -ForegroundColor Yellow "Failed to get dependencies"
        }

        $SWidth, $SHeight = Get-Resolution
        $IWidth, $IHeight = $SWidth, $SHeight

        if (-not $NoCrop) {
            $Image = Get-Image -FilePath $imagePath
            $IWidth = $Image.Width
            $IHeight = $Image.Height
            $IWidth, $IHeight = Get-CropRectangle -MonitorWidth $SWidth -MonitorHeight $SHeight -ImageWidth $IWidth -ImageHeight $IHeight
            $Image.Crop([SixLabors.ImageSharp.Rectangle]::new(0, 0, $IWidth, $IHeight))

            $SSmallSide = [math]::Min($SWidth, $SHeight)
            $ISmallSide = [math]::Min($IWidth, $IHeight)
            if ($ISmallSide -lt $SSmallSide) {
                $mult = $SSmallSide / $ISmallSide
                $Image.Resize($mult * 105)
                $Image.GaussianSharpen($SSmallSide / $ISmallSide / 1.5)
            }

            $IWidth = $Image.Width
            $IHeight = $Image.Height
            
            Save-Image -Image $Image -FilePath $imagePath
            Write-SafeHost "Image cropped to $IWidth $IHeight"
        }

        try {
            if (-not $NoText -and $title -is [string] -and -not ([System.String]::IsNullOrEmpty($title))) {
                Write-SafeHost "Setting title to $title"
                $fontSize = [math]::Floor($IHeight / 26)
                $padding = [math]::Min($IWidth / 30, $IHeight / 30)
                New-BGInfo {
                    New-BGInfoLabel -Name $title -Color White -FontSize $fontSize -FontFamilyName 'Segoe UI Semilight'
                } -FilePath $imagePath -ConfigurationDirectory $downloadDir -PositionX $padding -PositionY $padding -WallpaperFit Fill   
            }
            elseif ($hasDepend) {
                Set-DesktopWallpaper -Index 0 -Position Fill -WallpaperPath $imagePath
            }
            else {
                Set-Wallpaper -imagePath $imagePath
            }
            Write-SafeHost "Wallpaper set"            
        }
        catch {
            Write-SafeHost -BackgroundColor Red "Failed to set Wallpaper, trying again"
            Set-Wallpaper -imagePath $imagePath
        }
    }
    else {
        Write-SafeHost -ForegroundColor Red "Failed to get and update APOD"
    }
}

# Create new scheduled task to update bg image on every login
function New-SchedTask {
    if ([System.String]::IsNullOrEmpty($taskName)) {
        Write-SafeHost -ForegroundColor Red "Taskname cannot be empty!"
        return $false
    }

    if ($RunHidden) {
        if (Get-UnzipFile -DownloadUrl $RunHiddenURL -DestinationPath $downloadDir -FileName $RunHiddenName ) {
            $RunHiddenPath = Join-Path -Path $downloadDir -ChildPath $RunHiddenName
            $taskAction = New-ScheduledTaskAction -Execute $RunHiddenPath -Argument "$($PSCommandPath) -Update -NoAdmin -Silent"
        }
        else {
            Write-SafeHost -ForegroundColor Red "Failed to get RunHidden.exe, fallingback to just Powershell"
            $RunHidden = $false
        }
    }

    if (-not $RunHidden) {
        $taskAction = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File $($PSCommandPath) -Update -NoAdmin -Silent"
    }
    
    $taskTrigger = New-ScheduledTaskTrigger -AtLogOn
    $taskSettings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    $taskPrincipal = New-ScheduledTaskPrincipal -UserId ($env:USERDOMAIN + "\" + $env:USERNAME)
    $taskDescription = "Automatically download the latest Nasa APOD image and set it as the windows background image. Waits for an internet connection."
    $task = New-ScheduledTask -Action $taskAction -Trigger $taskTrigger -Settings $taskSettings -Principal $taskPrincipal -Description $taskDescription

    # Register task
    Register-ScheduledTask -TaskName $taskName -TaskPath "\" -InputObject $task
}

function Remove-SchedTask {
    if (-not (Get-ScheduledTask -TaskName $taskName)) {
        return $false
    }

    Unregister-ScheduledTask -TaskName $taskName -TaskPath "\" -Confirm:$false
    return $true
}

######### MAIN ##########

if ($Update) {
    Update
}
else {
    $confirmation = $null
    $done = $false

    $_ = Get-AdminRights
    $_ = Install-Dependencies

    if (-not $NoMove -and -not (Get-ScriptInDir -TargetDirectory $downloadDir)) {
        if ($Force) {
            $confirmation = 'y'
        }
        elseif (-not $Silent) {
            $confirmation = Read-Host "Script is not in the default location. Move it? [y/N]"
        }
        if ([System.String]::IsNullOrEmpty($confirmation)) { $confirmation = 'n' }
        if ($confirmation -eq 'y') {
            $_ = Move-ScriptAndRelaunch -TargetDirectory $downloadDir
        }
    }

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        if ($Force) {
            $confirmation = 'y'
        }
        else {
            if (-not $Silent) {
                $confirmation = Read-Host "Scheduled task already exists. Delete it? [y/N]"
            }
            if ([System.String]::IsNullOrEmpty($confirmation)) { $confirmation = 'n' }
        }

        if ($confirmation -eq 'y') {
            if (Remove-SchedTask) {
                if (-not $Silent) {
                    Write-SafeHost -ForegroundColor Green "[SUCCESS] Task has been deleted!"
                    Start-Sleep -Seconds 1
                }
            }
            else {
                if (-not $Silent) {
                    Write-SafeHost -ForegroundColor Red "[ERROR] Error while deleting task: $($Error[0].InnerException.Message)"
                    Read-Host
                }
            }
        }
        $done = $true
    }

    if ($Force) {
        $done = $false
    }

    if (-not $done) {
        if ($Force) {
            $confirmation = 'y'
        }
        elseif (-not $Silent) {
            $confirmation = Read-Host "Scheduled task does not exists. Create it? [y/N]"
        }
        if ([System.String]::IsNullOrEmpty($confirmation)) { $confirmation = 'n' }

        if ($confirmation -eq 'y') {
            if (New-SchedTask) {
                if (-not $Silent) {
                    Write-SafeHost -ForegroundColor Green "[SUCCESS] Task has been created!"
                    Start-Sleep -Seconds 1
                    Write-SafeHost ""
                    $confirm = Read-Host "Run task now? [y/N]"
                    if ([System.String]::IsNullOrEmpty($confirm)) { $confirm = 'n' }
                }

                if ($Force -or $confirm -eq 'y') {
                    Update
                }
            }
            else {
                if (-not $Silent) {
                    Write-SafeHost -ForegroundColor Red "[ERROR] Error while creating task: $($Error[0].InnerException.Message)"
                    Read-Host
                }
            }
        }
    }
    return 
    
}

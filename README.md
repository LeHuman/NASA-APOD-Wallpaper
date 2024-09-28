<!-- TITLE: NASA APOD Wallpaper -->
<!-- KEYWORDS: Wallpaper  -->
<!-- LANGUAGES: PowerShell -->

# NASA APOD Wallpaper

[About](#about) - [Usage](#usage) - [Related](#related) - [License](#license)

## Status

<!-- STATUS -->
**`Maintained`**

## About
<!-- DESCRIPTION START -->
PowerShell Script to automatically download the latest NASA APOD image and set it as the windows background image.
<!-- DESCRIPTION END -->

### Why

I have a similar app on my phone and wanted the same on my work laptop. I thought the original script would be enough, but I decided to fork, as I got carried away with my 'small' changes.

## Usage

> [!NOTE]
> Only works on Windows and tested on Windows 10+

The easiest way to use is with the following command.

```ps1
.\Set-NasaApodWallpaper.ps1 -Force -Silent
```

With `-Force` The script will move itself to a folder in pictures (where the pictures will be downloaded to), and set a scheduled task to trigger whenever anyone logs in using [RunHidden](https://github.com/LesFerch/RunHidden). The script waits for an internet connection once started and has a set task limit of 5 mins.

> [!IMPORTANT]
> A UAC prompt will show. This is for installing the dependencies and for setting up the scheduled Task. If you just want a task and nothing else installed, use `-NoAdmin`.

Remove `-Silent` if you wish to see the console as it goes.

If you are okay with a messy command, run the following to download the script and then run it.

```ps1
$scriptUrl = 'https://raw.githubusercontent.com/LeHuman/NASA-APOD-Wallpaper/main/Set-NasaApodWallpaper.ps1'; Set-Location $env:TEMP; $tempScript = ".\Set-NasaApodWallpaper.ps1"; Invoke-WebRequest -Uri $scriptUrl -OutFile $tempScript; powershell.exe -executionpolicy bypass .\Set-NasaApodWallpaper.ps1 -Force -Silent
```

### Requirements

- [Powershell](https://microsoft.com/PowerShell) >= 5.1
- Optionally installed by script, or install yourself I guess.
  - [NuGet](https://www.nuget.org/downloads) >= 2.8.5.201
  - [PowerBGInfo](https://github.com/EvotecIT/PowerBGInfo) >= v0.0.3

> [!WARNING]
> Does not work under [VSCode with Powershell extension](https://github.com/EvotecIT/PowerBGInfo?tab=readme-ov-file#known-issues) when *not* using `-NoAdmin`

### Parameters

- `-Update` [\<SwitchParameter\>]\
Only update the wallpaper. Don't ask for input unless needed. Will not setup a task.
- `-NoText` [\<SwitchParameter\>]\
Do not add any text to the wallpaper.
- `-NoCrop` [\<SwitchParameter\>]\
Do not crop the wallpaper to monitor ratio, text may not be visible.
- `-NoMove` [\<SwitchParameter\>]\
Do not move the script to the APOD folder. Make sure the script is not moved after setup.
- `-NoAdmin` [\<SwitchParameter\>]\
Do not install anything requring admin rights, reducing functionality. UAC prompt will still occur for task setup.
- `-NoDownload` [\<SwitchParameter\>]\
Do not download or install anything extra. Will set other flags as needed, reducing functionality.
        A console will appear when scheduled task is run.
- `-Silent` [\<SwitchParameter\>]\
Run silently, will not wait for user input or output anything. Will attempt to install anything necessary.\
May still trigger UAC prompt for setup.\
Consoles will appear on first set-up to show installation of dependencies.
- `-Force` [\<SwitchParameter\>]\
Force actions where possible, such as redownloading the APOD

### Print Help

```ps1
Get-Help .\Set-NasaApodWallpaper.ps1 -detailed
```

## Related

- EvotecIT/[PowerBGInfo](https://github.com/EvotecIT/PowerBGInfo)
- ex0tiq/[NASA-APOD-Wallpaper-Updater](https://github.com/ex0tiq/NASA-APOD-Wallpaper-Updater)
- LesFerch/[RunHidden](https://github.com/LesFerch/RunHidden)

## License

MIT license

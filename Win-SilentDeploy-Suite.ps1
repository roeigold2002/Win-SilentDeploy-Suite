# --- בדיקת הרשאות מנהל ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator!" -ForegroundColor Red
    exit
}

# --- הגדרות נתיבים ---
$BaseUrl = "https://github.com/roeigold2002/Win-SilentDeploy-Suite/releases/download/Softwares"
$InstallPath = "C:\OfflineInstalls"

if (-not (Test-Path $InstallPath)) { 
    New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null 
}

# --- רשימת אפליקציות ופרמטרים (מעודכן לקבצים שלך) ---
$Apps = @(
    @{ Name="Visual C++"; File="VC_redist.x64.exe"; Args="/quiet /norestart" },
    @{ Name="Google Chrome"; File="ChromeStandaloneSetup64.exe"; Args="/silent /install" },
    @{ Name="7-Zip"; File="7z2401-x64.exe"; Args="/S" },
    @{ Name="Notepad++"; File="npp.8.8.8.Installer.x64.exe"; Args="/S" },
    @{ Name="Python 3"; File="python-3.13.4-amd64.exe"; Args="/quiet InstallAllUsers=1 PrependPath=1" },
    @{ Name="Telegram"; File="tsetup-x64.6.5.1.exe"; Args="/VERYSILENT /ALLUSERS" },
    @{ Name="WinRAR"; File="winrar-x64-701.exe"; Args="/S" },
    @{ Name="OpenOffice"; File="Apache_OpenOffice_4.1.16_Win_x86_install_en-US.exe"; Args="/S" },
    @{ Name="Acrobat Reader"; File="Reader_en_install.exe"; Args="/sAll /rs" },
    @{ Name="Tor Browser"; File="tor-browser-windows-x86_64-portable-15.0.5.exe"; Args="SPECIAL" }
)

# --- שלב 1: הורדה מה-Repository שלך ---
Write-Host "--- Starting Downloads from GitHub Releases ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    $DownloadUrl = "$BaseUrl/$($App.File)"
    Write-Host "Downloading $($App.Name)... " -NoNewline
    try {
        Invoke-WebRequest -Uri $DownloadUrl -OutFile "$InstallPath\$($App.File)" -ErrorAction Stop
        Write-Host "Done." -ForegroundColor Green
    } catch {
        Write-Host "Failed! Check if file exists in Release." -ForegroundColor Red
    }
}

# --- שלב 2: התקנה שקטה ---
Write-Host "`n--- Starting Silent Installations ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    $LocalFile = "$InstallPath\$($App.File)"
    
    if (Test-Path $LocalFile) {
        if ($App.Name -eq "Tor Browser") {
            # טיפול מיוחד ב-Tor
            Write-Host "Extracting Tor Browser to C:\Tor... " -NoNewline
            $TorDest = "C:\Tor"
            if (-not (Test-Path $TorDest)) { New-Item -Path $TorDest -ItemType Directory -Force | Out-Null }
            
            Start-Process -FilePath $LocalFile -ArgumentList "/S /D=$TorDest" -Wait
            
            # הרשאות וקיצור דרך
            $Acl = Get-Acl $TorDest
            $Acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("Users","Modify","ContainerInherit,ObjectInherit","None","Allow")))
            Set-Acl $TorDest $Acl
            
            $WshShell = New-Object -ComObject WScript.Shell
            $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Tor Browser.lnk")
            $Shortcut.TargetPath = "$TorDest\Browser\firefox.exe"
            $Shortcut.Save()
            Write-Host "Done." -ForegroundColor Green
        } else {
            # התקנה רגילה
            Write-Host "Installing $($App.Name)... " -NoNewline
            $Process = Start-Process -FilePath $LocalFile -ArgumentList $App.Args -Wait -PassThru
            Write-Host "Done (Exit Code: $($Process.ExitCode))." -ForegroundColor Green
        }
    }
}

# --- שלב 3: ניקוי וריסטארט ---
Write-Host "`nCleaning up installation files... " -NoNewline
Remove-Item -Path $InstallPath -Recurse -Force
Write-Host "Done." -ForegroundColor Green

Write-Host "`nAll tasks completed. System will restart in 15 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 15
Restart-Computer -Force

# --- בדיקת הרשאות מנהל ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator!" -ForegroundColor Red
    exit
}

# הגדרות נתיבים ולינקים
$BaseUrl = "https://github.com/roeigold2002/Win-SilentDeploy-Suite/releases/download/Softwares"
$InstallPath = "C:\OfflineInstalls"
if (-not (Test-Path $InstallPath)) { New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null }

# --- רשימת אפליקציות (ללא אקרובט - הוא יותקן בסוף) ---
$Apps = @(
    @{ Name="Visual C++"; File="VC_redist.x64.exe"; Args="/quiet /norestart" },
    @{ Name="Google Chrome"; File="ChromeStandaloneSetup64.exe"; Args="/silent /install" },
    @{ Name="7-Zip"; File="7z2401-x64.exe"; Args="/S" },
    @{ Name="Notepad++"; File="npp.8.8.8.Installer.x64.exe"; Args="/S" },
    @{ Name="Python 3"; File="python-3.13.4-amd64.exe"; Args="/quiet InstallAllUsers=1 PrependPath=1" },
    @{ Name="Telegram"; File="tsetup-x64.6.5.1.exe"; Args="/VERYSILENT /ALLUSERS" },
    @{ Name="WinRAR"; File="winrar-x64-701.exe"; Args="/S" },
    @{ Name="OpenOffice"; File="Apache_OpenOffice_4.1.16_Win_x86_install_en-US.exe"; Args="/S" },
    @{ Name="Tor Browser"; File="tor-browser-windows-x86_64-portable-15.0.5.exe"; Args="SPECIAL" }
)

# --- שלב 1: הורדת כל הקבצים (כולל אקרובט) ---
Write-Host "--- Starting Downloads from GitHub ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    Write-Host "Downloading $($App.Name)... " -NoNewline
    Invoke-WebRequest -Uri "$BaseUrl/$($App.File)" -OutFile "$InstallPath\$($App.File)" -ErrorAction SilentlyContinue
    Write-Host "Done." -ForegroundColor Green
}
Write-Host "Downloading Acrobat Reader (Final Item)... " -NoNewline
Invoke-WebRequest -Uri "$BaseUrl/Reader_en_install.exe" -OutFile "$InstallPath\Reader_en_install.exe" -ErrorAction SilentlyContinue
Write-Host "Done." -ForegroundColor Green

# --- שלב 2: התקנת האפליקציות הראשיות ---
Write-Host "`n--- Installing Primary Applications ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    $LocalFile = "$InstallPath\$($App.File)"
    if (Test-Path $LocalFile) {
        if ($App.Name -eq "Tor Browser") {
            Write-Host "Extracting Tor to C:\Tor... " -NoNewline
            $TorDest = "C:\Tor"
            if (-not (Test-Path $TorDest)) { New-Item -Path $TorDest -ItemType Directory -Force | Out-Null }
            Start-Process -FilePath $LocalFile -ArgumentList "/S /D=$TorDest" -Wait
            $Acl = Get-Acl $TorDest; $Acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("Users","Modify","ContainerInherit,ObjectInherit","None","Allow"))); Set-Acl $TorDest $Acl
            $WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Tor Browser.lnk"); $Shortcut.TargetPath = "C:\Tor\Browser\firefox.exe"; $Shortcut.Save()
            Write-Host "Done." -ForegroundColor Green
        } else {
            Write-Host "Installing $($App.Name)... " -NoNewline
            Start-Process -FilePath $LocalFile -ArgumentList $App.Args -Wait
            Write-Host "Done." -ForegroundColor Green
        }
    }
}

# --- שלב 3: התקנת אקרובט בסוף עם הגנת זמן ---
Write-Host "`n--- Final Step: Installing Acrobat Reader ---" -ForegroundColor Cyan
$AcroFile = "$InstallPath\Reader_en_install.exe"
if (Test-Path $AcroFile) {
    Write-Host "Launching Acrobat (Silent Flags Applied)... " -NoNewline
    # שימוש בדגלים מתקדמים למניעת חלונות סיום
    $AcroProcess = Start-Process -FilePath $AcroFile -ArgumentList "/sAll /sPB /rs /msi EULA_ACCEPT=YES" -PassThru
    
    # המתנה של מקסימום 60 שניות - אם לא נסגר, ממשיכים הלאה
    $Timer = 0
    while (-not $AcroProcess.HasExited -and $Timer -lt 60) {
        Start-Sleep -Seconds 2
        $Timer += 2
    }
    Write-Host "Moving to cleanup." -ForegroundColor Green
}

# --- ניקוי וסיום ---
Write-Host "`nCleaning up... " -NoNewline
Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Done." -ForegroundColor Green

Write-Host "`nAll software installed. System will restart in 15 seconds." -ForegroundColor Yellow
Write-Host "If Acrobat is still open, the restart will close it automatically."
Start-Sleep -Seconds 15
Restart-Computer -Force

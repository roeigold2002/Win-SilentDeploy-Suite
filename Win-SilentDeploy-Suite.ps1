# הרצה כמנהל מערכת
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator!" -ForegroundColor Red
    exit
}

$InstallPath = "C:\OfflineInstalls"
if (-not (Test-Path $InstallPath)) { New-Item -Path $InstallPath -ItemType Directory -Force | Out-Null }

# --- רשימת התוכנות והגדרות ---
$Apps = @(
    @{ Name="Visual C++ 2015-2022"; Url="https://aka.ms/vs/17/release/vc_redist.x64.exe"; File="vc_redist.exe"; Args="/quiet /norestart" },
    @{ Name="DirectX Redist"; Url="https://download.microsoft.com/download/1/7/1/1718CCC4-6315-4D8E-9543-8E28A4E18C4C/directx_Jun2010_redist.exe"; File="directx_redist.exe"; Args="/Q /T:$InstallPath\DX" },
    @{ Name="Google Chrome"; Url="https://dl.google.com/chrome/install/standalonesetup64.exe"; File="chrome_setup.exe"; Args="/silent /install" },
    @{ Name="7-Zip"; Url="https://www.7-zip.org/a/7z2401-x64.exe"; File="7z_setup.exe"; Args="/S" },
    @{ Name="Notepad++"; Url="https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.8.8/npp.8.8.8.Installer.x64.exe"; File="npp_setup.exe"; Args="/S" },
    @{ Name="Python 3"; Url="https://www.python.org/ftp/python/3.13.4/python-3.13.4-amd64.exe"; File="python_setup.exe"; Args="/quiet InstallAllUsers=1 PrependPath=1" },
    @{ Name="Acrobat Reader"; Url="https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2400320112/AcrobatRDC2400320112_en_US.exe"; File="acrobat_setup.exe"; Args="/sAll /rs" },
    @{ Name="Telegram"; Url="https://tdesktop.com/win64/current?setup=1"; File="telegram_setup.exe"; Args="/VERYSILENT /SUPPRESSMSGBOXES /ALLUSERS" },
    @{ Name="WinRAR"; Url="https://www.rarlab.com/rar/winrar-x64-701.exe"; File="winrar_setup.exe"; Args="/S" },
    @{ Name="OpenOffice"; Url="https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.16/binaries/en-US/Apache_OpenOffice_4.1.16_Win_x86_install_en-US.exe/download"; File="openoffice_setup.exe"; Args="/S" },
    @{ Name="Tor Browser"; Url="https://www.torproject.org/dist/torbrowser/14.0.5/tor-browser-windows-x86_64-portable-14.0.5.exe"; File="tor_setup.exe"; Args="SPECIAL" }
)

# --- 1. הורדה ---
Write-Host "--- Starting Downloads ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    Write-Host "Downloading $($App.Name)... " -NoNewline
    try {
        Invoke-WebRequest -Uri $App.Url -OutFile "$InstallPath\$($App.File)" -ErrorAction Stop
        Write-Host "Done." -ForegroundColor Green
    } catch { Write-Host "Failed!" -ForegroundColor Red }
}

# --- 2. התקנת Runtimes (DirectX & VC) ---
Write-Host "`n--- Installing System Runtimes ---" -ForegroundColor Cyan
# DirectX
if (Test-Path "$InstallPath\directx_redist.exe") {
    New-Item -Path "$InstallPath\DX" -ItemType Directory -Force | Out-Null
    Start-Process -FilePath "$InstallPath\directx_redist.exe" -ArgumentList "/Q /T:$InstallPath\DX" -Wait
    Start-Process -FilePath "$InstallPath\DX\dxsetup.exe" -ArgumentList "/silent" -Wait
}
# Visual C++
Start-Process -FilePath "$InstallPath\vc_redist.exe" -ArgumentList "/quiet /norestart" -Wait

# --- 3. התקנת אפליקציות ---
Write-Host "`n--- Installing Applications ---" -ForegroundColor Cyan
foreach ($App in $Apps) {
    if ($App.Name -match "DirectX" -or $App.Name -match "Visual C++" -or $App.Name -match "Tor Browser") { continue }
    if (Test-Path "$InstallPath\$($App.File)") {
        Write-Host "Installing $($App.Name)... " -NoNewline
        Start-Process -FilePath "$InstallPath\$($App.File)" -ArgumentList $App.Args -Wait
        Write-Host "Done." -ForegroundColor Green
    }
}

# --- 4. טיפול ב-Tor ---
$TorDest = "C:\Tor"
if (Test-Path "$InstallPath\tor_setup.exe") {
    Write-Host "Configuring Tor Browser for all users... " -NoNewline
    if (-not (Test-Path $TorDest)) { New-Item -Path $TorDest -ItemType Directory -Force | Out-Null }
    Start-Process -FilePath "$InstallPath\tor_setup.exe" -ArgumentList "/S /D=$TorDest" -Wait
    
    $Acl = Get-Acl $TorDest
    $Acl.SetAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule("Users","Modify","ContainerInherit,ObjectInherit","None","Allow")))
    Set-Acl $TorDest $Acl

    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("C:\Users\Public\Desktop\Tor Browser.lnk")
    $Shortcut.TargetPath = "$TorDest\Browser\firefox.exe"; $Shortcut.Save()
    Write-Host "Done." -ForegroundColor Green
}

# --- 5. ניקוי סופי (Cleanup) ---
Write-Host "`n--- Cleaning up installation files ---" -ForegroundColor Cyan
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force
    Write-Host "Temp files removed." -ForegroundColor Green
}

# --- ריסטארט ---
Write-Host "`nSetup complete. Restarting in 15 seconds..." -ForegroundColor Yellow
Start-Sleep -Seconds 15
Restart-Computer -Force

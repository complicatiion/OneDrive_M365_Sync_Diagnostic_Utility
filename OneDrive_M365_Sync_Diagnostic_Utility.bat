@echo off
setlocal EnableExtensions
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS_EXE%" set "PS_EXE=powershell.exe"

title OneDrive and M365 Sync Diagnostic Utility
color 0B
chcp 65001 >nul

net session >nul 2>&1
if errorlevel 1 (
  set "ISADMIN=0"
) else (
  set "ISADMIN=1"
)

set "REPORTROOT=%USERPROFILE%\Desktop\OneDriveSyncReports"
if not exist "%REPORTROOT%" md "%REPORTROOT%" >nul 2>&1
set "GPDIR=%SystemDrive%\Temp"
if not exist "%GPDIR%" md "%GPDIR%" >nul 2>&1

:MAIN
cls
echo ============================================================
echo.
echo  OneDrive and M365 Sync Diagnostic Utility by complicatiion
echo.
echo ============================================================
echo.
if "%ISADMIN%"=="1" (
  echo Admin Status: YES
) else (
  echo Admin Status: NO
)
echo Report Folder : %REPORTROOT%
echo GPResult Folder: %GPDIR%
echo.
echo [1] Quick analysis OneDrive / SharePoint / M365 sync
echo [2] Generate GPResult HTML report
echo [3] Check OneDrive / KFM / SharePoint policies
echo [4] Check OneDrive client, version, process and services
echo [5] Check sync accounts, known folders and sync paths
echo [6] Check OneDrive / SharePoint / Office event logs
echo [7] Check Office integration, identity and cache environment
echo [8] Start or restart OneDrive client
echo [9] Open OneDrive settings locations
echo [A] Create full report
echo [B] Open report folder
echo [0] Exit
echo.
set "CHO="
set /p CHO="Selection: "

if "%CHO%"=="1" goto QUICK
if "%CHO%"=="2" goto GPRESULT
if "%CHO%"=="3" goto POLICIES
if "%CHO%"=="4" goto CLIENT
if "%CHO%"=="5" goto SYNCSTATE
if "%CHO%"=="6" goto EVENTS
if "%CHO%"=="7" goto OFFICEENV
if "%CHO%"=="8" goto STARTOD
if "%CHO%"=="9" goto OPENPATHS
if /I "%CHO%"=="A" goto REPORT
if /I "%CHO%"=="B" goto OPENFOLDER
if "%CHO%"=="0" goto END
goto MAIN

:QUICK
cls
echo ============================================================
echo Quick analysis OneDrive / SharePoint / M365 sync
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- Device and user context ---'; [pscustomobject]@{Computer=$env:COMPUTERNAME; User=$env:USERNAME; UserProfile=$env:USERPROFILE; AppData=$env:APPDATA; LocalAppData=$env:LOCALAPPDATA} | Format-List; Write-Host ''; Write-Host '--- OneDrive process ---'; $proc=Get-Process OneDrive -ErrorAction SilentlyContinue; if($proc){ $proc | Select-Object Name,Id,Responding,CPU,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}},StartTime | Format-Table -AutoSize } else { Write-Host 'OneDrive process is not running.' }; Write-Host ''; Write-Host '--- Installed client ---'; $paths=@($env:LOCALAPPDATA + '\Microsoft\OneDrive\OneDrive.exe',$env:ProgramFiles + '\Microsoft OneDrive\OneDrive.exe',$env:ProgramFilesx86 + '\Microsoft OneDrive\OneDrive.exe'); $hit=$null; foreach($p in $paths){ if(Test-Path $p){ $hit=Get-Item $p; break } }; if($hit){ [pscustomobject]@{Path=$hit.FullName; Version=$hit.VersionInfo.FileVersion; LastWriteTime=$hit.LastWriteTime} | Format-List } else { Write-Host 'OneDrive.exe not found in common paths.' }; Write-Host ''; Write-Host '--- Policies quick view ---'; $policyPath='HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'; if(Test-Path $policyPath){ Get-ItemProperty $policyPath | Select-Object KFMBlockOptIn,KFMBlockOptOut,KFMOptInWithWizard,KFMOptInNoWizard,SilentAccountConfig,FilesOnDemandEnabled,DisablePersonalSync,TenantAutoMount | Format-List } else { Write-Host 'HKLM OneDrive policy key not found.' }; Write-Host ''; Write-Host '--- Known Folder redirection status ---'; $usf='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; if(Test-Path $usf){ $r=Get-ItemProperty $usf; [pscustomobject]@{Desktop=$r.Desktop; Documents=$r.Personal; Pictures=$r.MyPictures} | Format-List } else { Write-Host 'User Shell Folders key not found.' }; Write-Host ''; Write-Host '--- Sync roots under user profile ---'; $roots=Get-ChildItem -Path $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'OneDrive*' } | Select-Object Name,FullName; if($roots){ $roots | Format-Table -AutoSize } else { Write-Host 'No OneDrive sync root folders found under user profile.' }"
echo.
pause
goto MAIN

:GPRESULT
cls
echo ============================================================
echo Generate GPResult HTML report
echo ============================================================
echo.
set "GPHTML=%GPDIR%\gp.html"
gpresult /h "%GPHTML%"
echo.
if exist "%GPHTML%" (
  echo GPResult created:
  echo %GPHTML%
) else (
  echo GPResult could not be created.
)
echo.
pause
goto MAIN

:POLICIES
cls
echo ============================================================
echo Check OneDrive / KFM / SharePoint policies
echo ============================================================
echo.
echo [Registry export]
reg query "HKLM\SOFTWARE\Policies\Microsoft\OneDrive"
echo.
reg query "HKCU\SOFTWARE\Policies\Microsoft\OneDrive"
echo.
echo [Interpreted policy view]
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; function Get-PolicyRows($root){ if(Test-Path $root){ $p=Get-ItemProperty $root; $rows=@(); $defs=@(@{Name='KFMBlockOptIn'; Meaning='Blocks moving known folders into OneDrive'; BadWhen=1},@{Name='KFMBlockOptOut'; Meaning='Prevents moving known folders back to the PC'; BadWhen=1},@{Name='KFMOptInWithWizard'; Meaning='Prompts KFM wizard for tenant'; BadWhen=$null},@{Name='KFMOptInNoWizard'; Meaning='Silently redirects known folders for tenant'; BadWhen=$null},@{Name='SilentAccountConfig'; Meaning='Silently signs in user with Windows credentials'; BadWhen=$null},@{Name='FilesOnDemandEnabled'; Meaning='Enables Files On-Demand'; BadWhen=0},@{Name='DisablePersonalSync'; Meaning='Blocks personal Microsoft account sync'; BadWhen=$null},@{Name='TenantAutoMount'; Meaning='Automatically mounts team site libraries'; BadWhen=$null}); foreach($d in $defs){ $exists=$null -ne $p.PSObject.Properties[$d.Name]; $val=if($exists){ $p.($d.Name) } else { $null }; $status='Not set'; if($exists){ $status='Configured' }; $assessment='Info'; if($exists -and $null -ne $d.BadWhen -and [string]$val -eq [string]$d.BadWhen){ $assessment='Potential blocker' }; $rows += [pscustomobject]@{Scope=$root; Policy=$d.Name; Value=$val; Status=$status; Assessment=$assessment; Meaning=$d.Meaning} }; $rows } }; $all=@(); $all += Get-PolicyRows 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'; $all += Get-PolicyRows 'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive'; if($all){ $all | Format-Table -AutoSize } else { Write-Host 'No OneDrive policy keys found in HKLM or HKCU.' }"
echo.
echo Notes:
echo - KFMBlockOptIn=1 blocks moving known folders into OneDrive.
echo - KFMBlockOptOut=1 prevents moving known folders back to the local PC.
echo - FilesOnDemandEnabled=0 can cause expected cloud behavior to be unavailable.
echo - TenantAutoMount and KFM settings often explain SharePoint library or Known Folder issues.
echo.
pause
goto MAIN

:CLIENT
cls
echo ============================================================
echo Check OneDrive client, version, process and services
echo ============================================================
echo.
echo [1] OneDrive process
tasklist /v | findstr /I "OneDrive.exe"
echo.
echo [2] Client details and related services
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- OneDrive executable ---'; $paths=@($env:LOCALAPPDATA + '\Microsoft\OneDrive\OneDrive.exe',$env:ProgramFiles + '\Microsoft OneDrive\OneDrive.exe',$env:ProgramFilesx86 + '\Microsoft OneDrive\OneDrive.exe'); $rows=@(); foreach($p in $paths){ if(Test-Path $p){ $i=Get-Item $p; $rows += [pscustomobject]@{Path=$i.FullName; Version=$i.VersionInfo.FileVersion; ProductVersion=$i.VersionInfo.ProductVersion; LastWriteTime=$i.LastWriteTime} } }; if($rows){ $rows | Format-Table -AutoSize } else { Write-Host 'OneDrive.exe not found.' }; Write-Host ''; Write-Host '--- Process ---'; $proc=Get-Process OneDrive -ErrorAction SilentlyContinue; if($proc){ $proc | Select-Object Name,Id,Responding,CPU,Handles,Threads,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}},StartTime | Format-Table -AutoSize } else { Write-Host 'OneDrive process is not running.' }; Write-Host ''; Write-Host '--- Services relevant for SharePoint/WebDAV ---'; Get-Service WebClient,BITS,lanmanworkstation -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType | Format-Table -AutoSize; Write-Host ''; Write-Host '--- Startup / scheduled tasks ---'; Get-ScheduledTask -TaskPath '\Microsoft\OneDrive\' -ErrorAction SilentlyContinue | Select-Object TaskName,State,TaskPath | Format-Table -AutoSize"
echo.
pause
goto MAIN

:SYNCSTATE
cls
echo ============================================================
echo Check sync accounts, known folders and sync paths
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- OneDrive account registry ---'; $acctRoot='HKCU:\Software\Microsoft\OneDrive\Accounts'; if(Test-Path $acctRoot){ $rows=@(); Get-ChildItem $acctRoot -ErrorAction SilentlyContinue | ForEach-Object { $p=Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue; $rows += [pscustomobject]@{Account=$_.PSChildName; UserEmail=$p.UserEmail; UserFolder=$p.UserFolder; TenantName=$p.TenantName; Configured=$true} }; if($rows){ $rows | Format-Table -AutoSize } else { Write-Host 'Account root exists but no accounts were enumerated.' } } else { Write-Host 'No OneDrive account registry root found for current user.' }; Write-Host ''; Write-Host '--- Known folders current target ---'; $usf='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; if(Test-Path $usf){ $r=Get-ItemProperty $usf; [pscustomobject]@{Desktop=$r.Desktop; Documents=$r.Personal; Pictures=$r.MyPictures} | Format-List } else { Write-Host 'User Shell Folders not found.' }; Write-Host ''; Write-Host '--- Sync roots under profile ---'; $roots=Get-ChildItem -Path $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'OneDrive*' } | Select-Object Name,FullName,LastWriteTime; if($roots){ $roots | Format-Table -AutoSize } else { Write-Host 'No OneDrive folders found under user profile.' }; Write-Host ''; Write-Host '--- SharePoint / OneDrive cache and logs ---'; $check=@($env:LOCALAPPDATA + '\Microsoft\OneDrive',$env:LOCALAPPDATA + '\Microsoft\Office\16.0\OfficeFileCache',$env:LOCALAPPDATA + '\Microsoft\Office\Spw',$env:LOCALAPPDATA + '\Microsoft\OneDrive\logs',$env:LOCALAPPDATA + '\Microsoft\OneDrive\settings'); $pathRows=foreach($p in $check){ [pscustomobject]@{Path=$p; Exists=([bool](Test-Path $p))} }; $pathRows | Format-Table -AutoSize"
echo.
echo Notes:
echo - Desktop, Documents and Pictures pointing to a OneDrive path usually indicate KFM is active.
echo - Missing account entries or missing UserFolder values often indicate client sign-in or profile corruption.
echo - Missing OfficeFileCache can point to Office integration not having initialized yet.
echo.
pause
goto MAIN

:EVENTS
cls
echo ============================================================
echo Check OneDrive / SharePoint / Office event logs
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- Application log (last 14 days) ---'; $events=Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-14)} -ErrorAction SilentlyContinue; $result=foreach($ev in $events){ $msg=''; try { $msg=[string]$ev.Message } catch { $msg='' }; if($ev.ProviderName -match 'OneDrive|Office|Microsoft Office Alerts|Application Error|Application Hang|WebClient' -or $msg -match 'OneDrive|SharePoint|groove|OfficeFileCache|WebDAV|WebClient'){ [pscustomobject]@{TimeCreated=$ev.TimeCreated; Id=$ev.Id; ProviderName=$ev.ProviderName; Level=$ev.LevelDisplayName; Message=(($msg -replace '\r?\n',' ').Trim())} } }; if($result){ $result | Select-Object -First 80 | Format-List } else { Write-Host 'No relevant Application log entries found.' }; Write-Host ''; Write-Host '--- OneDrive operational logs if available ---'; $logNames=@('Microsoft-OneDrive/Operational','Microsoft-Windows-OneDrive/Operational'); foreach($ln in $logNames){ if((Get-WinEvent -ListLog $ln -ErrorAction SilentlyContinue)){ Write-Host ('Log: ' + $ln); Get-WinEvent -LogName $ln -MaxEvents 30 -ErrorAction SilentlyContinue | Select-Object TimeCreated,Id,LevelDisplayName,ProviderName,@{N='Message';E={ try { ([string]$_.Message -replace '\r?\n',' ').Trim() } catch { '' } }} | Format-Table -AutoSize; Write-Host '' } }"
echo.
pause
goto MAIN

:OFFICEENV
cls
echo ============================================================
echo Check Office integration, identity and cache environment
echo ============================================================
echo.
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; Write-Host '--- Office identity keys ---'; $keys=@('HKCU:\Software\Microsoft\Office\16.0\Common\Identity','HKCU:\Software\Microsoft\Office\16.0\Common\Internet','HKCU:\Software\Microsoft\Office\16.0\Common\Licensing'); foreach($k in $keys){ if(Test-Path $k){ Write-Host ('Key: ' + $k); Get-ItemProperty $k | Format-List; Write-Host '' } else { Write-Host ('Missing: ' + $k) } }; Write-Host ''; Write-Host '--- Office and OneDrive paths ---'; $check=@($env:LOCALAPPDATA + '\Microsoft\Office\16.0\OfficeFileCache',$env:LOCALAPPDATA + '\Microsoft\Office',$env:APPDATA + '\Microsoft\Office',$env:LOCALAPPDATA + '\Microsoft\OneDrive',$env:LOCALAPPDATA + '\Packages\MSTeams_8wekyb3d8bbwe\LocalCache'); $rows=foreach($p in $check){ [pscustomobject]@{Path=$p; Exists=([bool](Test-Path $p))} }; $rows | Format-Table -AutoSize; Write-Host ''; Write-Host '--- Environment variables ---'; Get-ChildItem Env: | Where-Object { $_.Name -match 'ONEDRIVE|SHAREPOINT|APPDATA|PROFILE|TEMP|ODOPEN|GROOVE' } | Sort-Object Name | Format-Table -AutoSize"
echo.
echo Notes:
echo - OfficeFileCache issues can break coauthoring and file open/save sync behavior.
echo - Missing Office identity state can point to broken Office sign-in or token problems.
echo - SharePoint sync issues are often a mix of OneDrive client state, policy and Office cache state.
echo.
pause
goto MAIN

:STARTOD
cls
echo ============================================================
echo Start or restart OneDrive client
echo ============================================================
echo.
taskkill /IM OneDrive.exe /F >nul 2>&1
if exist "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe" (
  start "" "%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe"
  echo OneDrive started from LocalAppData.
) else if exist "%ProgramFiles%\Microsoft OneDrive\OneDrive.exe" (
  start "" "%ProgramFiles%\Microsoft OneDrive\OneDrive.exe"
  echo OneDrive started from Program Files.
) else if exist "%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe" (
  start "" "%ProgramFiles(x86)%\Microsoft OneDrive\OneDrive.exe"
  echo OneDrive started from Program Files x86.
) else (
  echo OneDrive executable not found in common locations.
)
echo.
pause
goto MAIN

:OPENPATHS
cls
echo ============================================================
echo Open OneDrive settings locations
echo ============================================================
echo.
if exist "%LOCALAPPDATA%\Microsoft\OneDrive" start "" explorer.exe "%LOCALAPPDATA%\Microsoft\OneDrive"
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache" start "" explorer.exe "%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache"
if exist "%USERPROFILE%" start "" explorer.exe "%USERPROFILE%"
echo Opened existing locations where possible.
echo.
pause
goto MAIN

:REPORT
cls
echo [*] Creating report...
echo.
set "STAMP=%DATE%_%TIME%"
set "STAMP=%STAMP:/=-%"
set "STAMP=%STAMP:\=-%"
set "STAMP=%STAMP::=-%"
set "STAMP=%STAMP:.=-%"
set "STAMP=%STAMP:,=-%"
set "STAMP=%STAMP: =0%"
set "OUTFILE=%REPORTROOT%\OneDrive_M365_Sync_Report_%STAMP%.txt"
set "GPHTML=%GPDIR%\gp.html"

> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo OneDrive and M365 Sync Diagnostic Report
>> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo Date: %DATE% %TIME%
>> "%OUTFILE%" echo Computer: %COMPUTERNAME%
>> "%OUTFILE%" echo User: %USERNAME%
>> "%OUTFILE%" echo Admin: %ISADMIN%
>> "%OUTFILE%" echo Report Folder: %REPORTROOT%
>> "%OUTFILE%" echo GPResult HTML: %GPHTML%
>> "%OUTFILE%" echo ============================================================
>> "%OUTFILE%" echo.
>> "%OUTFILE%" echo [1] GPResult
gpresult /h "%GPHTML%" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [2] Raw OneDrive policy registry
) >> "%OUTFILE%"
reg query "HKLM\SOFTWARE\Policies\Microsoft\OneDrive" >> "%OUTFILE%" 2>&1
reg query "HKCU\SOFTWARE\Policies\Microsoft\OneDrive" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [3] Interpreted OneDrive policy state
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; function Get-PolicyRows($root){ if(Test-Path $root){ $p=Get-ItemProperty $root; $rows=@(); $defs=@(@{Name='KFMBlockOptIn'; Meaning='Blocks moving known folders into OneDrive'; BadWhen=1},@{Name='KFMBlockOptOut'; Meaning='Prevents moving known folders back to the PC'; BadWhen=1},@{Name='KFMOptInWithWizard'; Meaning='Prompts KFM wizard for tenant'; BadWhen=$null},@{Name='KFMOptInNoWizard'; Meaning='Silently redirects known folders for tenant'; BadWhen=$null},@{Name='SilentAccountConfig'; Meaning='Silently signs in user with Windows credentials'; BadWhen=$null},@{Name='FilesOnDemandEnabled'; Meaning='Enables Files On-Demand'; BadWhen=0},@{Name='DisablePersonalSync'; Meaning='Blocks personal Microsoft account sync'; BadWhen=$null},@{Name='TenantAutoMount'; Meaning='Automatically mounts team site libraries'; BadWhen=$null}); foreach($d in $defs){ $exists=$null -ne $p.PSObject.Properties[$d.Name]; $val=if($exists){ $p.($d.Name) } else { $null }; $status='Not set'; if($exists){ $status='Configured' }; $assessment='Info'; if($exists -and $null -ne $d.BadWhen -and [string]$val -eq [string]$d.BadWhen){ $assessment='Potential blocker' }; [pscustomobject]@{Scope=$root; Policy=$d.Name; Value=$val; Status=$status; Assessment=$assessment; Meaning=$d.Meaning} } } }; $all=@(); $all += Get-PolicyRows 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'; $all += Get-PolicyRows 'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive'; if($all){ $all | Format-Table -AutoSize } else { 'No OneDrive policy keys found.' }" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [4] OneDrive client, processes and services
) >> "%OUTFILE%"
tasklist /v | findstr /I "OneDrive.exe" >> "%OUTFILE%" 2>&1
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $paths=@($env:LOCALAPPDATA + '\Microsoft\OneDrive\OneDrive.exe',$env:ProgramFiles + '\Microsoft OneDrive\OneDrive.exe',$env:ProgramFilesx86 + '\Microsoft OneDrive\OneDrive.exe'); $rows=@(); foreach($p in $paths){ if(Test-Path $p){ $i=Get-Item $p; $rows += [pscustomobject]@{Path=$i.FullName; Version=$i.VersionInfo.FileVersion; ProductVersion=$i.VersionInfo.ProductVersion; LastWriteTime=$i.LastWriteTime} } }; if($rows){ $rows | Format-Table -AutoSize } else { 'OneDrive.exe not found.' }; ''; $proc=Get-Process OneDrive -ErrorAction SilentlyContinue; if($proc){ $proc | Select-Object Name,Id,Responding,CPU,Handles,Threads,@{N='WorkingSetMB';E={[math]::Round($_.WorkingSet64/1MB,1)}},StartTime | Format-Table -AutoSize } else { 'OneDrive process is not running.' }; ''; Get-Service WebClient,BITS,lanmanworkstation -ErrorAction SilentlyContinue | Select-Object Name,DisplayName,Status,StartType | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [5] Sync accounts, known folders and cache paths
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $acctRoot='HKCU:\Software\Microsoft\OneDrive\Accounts'; if(Test-Path $acctRoot){ Get-ChildItem $acctRoot -ErrorAction SilentlyContinue | ForEach-Object { $p=Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue; [pscustomobject]@{Account=$_.PSChildName; UserEmail=$p.UserEmail; UserFolder=$p.UserFolder; TenantName=$p.TenantName} } | Format-Table -AutoSize } else { 'No OneDrive account registry root found.' }; ''; $usf='HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders'; if(Test-Path $usf){ $r=Get-ItemProperty $usf; [pscustomobject]@{Desktop=$r.Desktop; Documents=$r.Personal; Pictures=$r.MyPictures} | Format-List }; ''; $roots=Get-ChildItem -Path $env:USERPROFILE -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like 'OneDrive*' } | Select-Object Name,FullName,LastWriteTime; if($roots){ $roots | Format-Table -AutoSize } else { 'No OneDrive folders found under user profile.' }; ''; $check=@($env:LOCALAPPDATA + '\Microsoft\OneDrive',$env:LOCALAPPDATA + '\Microsoft\Office\16.0\OfficeFileCache',$env:LOCALAPPDATA + '\Microsoft\Office\Spw',$env:LOCALAPPDATA + '\Microsoft\OneDrive\logs',$env:LOCALAPPDATA + '\Microsoft\OneDrive\settings'); $pathRows=foreach($p in $check){ [pscustomobject]@{Path=$p; Exists=([bool](Test-Path $p))} }; $pathRows | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [6] Office identity and environment
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $keys=@('HKCU:\Software\Microsoft\Office\16.0\Common\Identity','HKCU:\Software\Microsoft\Office\16.0\Common\Internet','HKCU:\Software\Microsoft\Office\16.0\Common\Licensing'); foreach($k in $keys){ if(Test-Path $k){ 'Key: ' + $k; Get-ItemProperty $k | Format-List; '' } else { 'Missing: ' + $k } }; ''; Get-ChildItem Env: | Where-Object { $_.Name -match 'ONEDRIVE|SHAREPOINT|APPDATA|PROFILE|TEMP|ODOPEN|GROOVE' } | Sort-Object Name | Format-Table -AutoSize" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [7] Event logs
) >> "%OUTFILE%"
"%PS_EXE%" -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='SilentlyContinue'; $events=Get-WinEvent -FilterHashtable @{LogName='Application'; StartTime=(Get-Date).AddDays(-14)} -ErrorAction SilentlyContinue; $result=foreach($ev in $events){ $msg=''; try { $msg=[string]$ev.Message } catch { $msg='' }; if($ev.ProviderName -match 'OneDrive|Office|Microsoft Office Alerts|Application Error|Application Hang|WebClient' -or $msg -match 'OneDrive|SharePoint|groove|OfficeFileCache|WebDAV|WebClient'){ [pscustomobject]@{TimeCreated=$ev.TimeCreated; Id=$ev.Id; ProviderName=$ev.ProviderName; Level=$ev.LevelDisplayName; Message=(($msg -replace '\r?\n',' ').Trim())} } }; if($result){ $result | Select-Object -First 80 | Format-List } else { 'No relevant Application log entries found.' }" >> "%OUTFILE%" 2>&1
(
  echo.
  echo [8] Interpretation
  echo - KFMBlockOptIn=1 blocks Known Folder Move into OneDrive.
  echo - KFMBlockOptOut=1 prevents moving known folders back to the local PC.
  echo - Missing OneDrive account entries often point to client sign-in, user profile or token issues.
  echo - Missing OfficeFileCache or broken Office identity state can affect SharePoint and M365 open/save sync behavior.
  echo - WebClient service state can matter for legacy SharePoint and WebDAV-related behavior.
) >> "%OUTFILE%"

echo Report created:
echo %OUTFILE%
echo.
pause
goto MAIN

:OPENFOLDER
start "" explorer.exe "%REPORTROOT%"
goto MAIN

:END
endlocal
exit /b 0

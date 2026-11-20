@echo off
title Cong cu Ho tro & Sua loi Cap Nhat Windows 11
color 0A

:: Kiem tra quyen Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [LOI] Vui long chuot phai vao file nay va chon "Run as administrator"!
    pause
    exit /b
)

echo ============================================================================
echo   THONG TIN NGON NGU WINDOWS HIEN TAI (DE TAI DUNG FILE ISO)
echo ============================================================================
powershell -NoProfile -Command "$r = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language'; $inst = $r.InstallLanguage; $def = $r.Default; $instName = try { [System.Globalization.CultureInfo]::GetCultureInfo([Convert]::ToInt32($inst,16)).DisplayName } catch { 'Khong xac dinh' }; $defName = try { [System.Globalization.CultureInfo]::GetCultureInfo([Convert]::ToInt32($def,16)).DisplayName } catch { 'Khong xac dinh' }; Write-Host ('  - Ngon ngu GOC cai dat (InstallLanguage): ' + $inst + ' -> ' + $instName); Write-Host ('  - Ngon ngu HIEN THI (Default)           : ' + $def + ' -> ' + $defName); if ($inst -eq '0809') { Write-Host '  => Luu y: Ma 0809 tuong ung voi ban ISO [English International] (neu trang /fr-fr/ khong co, hay doi URL thanh /en-us/).' -ForegroundColor Yellow }"
echo   (Ngon ngu file ISO bat buoc phai trung voi Ngon ngu GOC de giu lai ung dung)
echo.

echo ============================================================================
echo   LUA CHON CHE DO CHAY
echo ============================================================================
echo   [Y] Yes - Da update thanh cong, CHI XOA CAC FILE UPDATE de tiet kiem bo nho
echo             (Xoa $WINDOWS.~BT, $WINDOWS.~WS, Windows.old, bo nho dem Update)
echo             * Luu y: Xoa Windows.old se khong quay lai ban Windows cu duoc nua.
echo   [N] No  - CHUA update duoc, chay day du cac buoc SUA LOI (WinRE, DISM, SFC...)
echo.
choice /C YN /N /M "Chon [Y] de don dep tiet kiem bo nho, hoac [N] de sua loi update (Y/N): "
if errorlevel 2 goto :FIX_MODE
if errorlevel 1 goto :CLEANUP_ONLY

:CLEANUP_ONLY
echo.
echo ============================================================================
echo   CHE DO DON DEP SAU UPDATE (TIET KIEM BO NHO)
echo ============================================================================
call :StopUpdateServices

call :DeleteFolder "C:\$WINDOWS.~BT"
call :DeleteFolder "C:\$WINDOWS.~WS"
if exist "C:\Windows.old" (
    echo Dang xoa C:\Windows.old bang DISM ...
    DISM /Online /Remove-OSUninstall >nul 2>&1
)
if exist "C:\Windows.old" (
    echo Dang xoa C:\Windows.old bang Disk Cleanup ^(cleanmgr^)...
    reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations" /v StateFlags0099 /t REG_DWORD /d 2 /f >nul 2>&1
    cleanmgr /sagerun:99
)
if exist "C:\Windows.old" (
    echo [LUU Y] Mot so file trong C:\Windows.old van bi he thong khoa.
    echo Hay mo "cleanmgr" -^> "Clean up system files" -^> chon "Previous Windows installation(s)" de xoa not.
)
call :DeleteFolder "C:\Windows\SoftwareDistribution.old"
call :DeleteFolder "C:\Windows\System32\catroot2.old"

if exist "C:\Windows\SoftwareDistribution\Download" (
    echo Dang xoa bo nho dem tai ve trong SoftwareDistribution\Download...
    del /F /S /Q "C:\Windows\SoftwareDistribution\Download\*" >nul 2>&1
    for /d %%D in ("C:\Windows\SoftwareDistribution\Download\*") do rmdir /S /Q "%%D" >nul 2>&1
)
del /F /Q "C:\ProgramData\Microsoft\Network\Downloader\qmgr*.dat" >nul 2>&1

call :StartUpdateServices

echo.
echo ============================================================================
echo   DA DON DEP XONG CAC FILE UPDATE!
echo ============================================================================
pause
exit /b

:FIX_MODE
echo.
echo ============================================================================
echo   BUOC 1: DUNG CAC DICH VU WINDOWS UPDATE
echo ============================================================================
call :StopUpdateServices

echo.
echo ============================================================================
echo   BUOC 2: XOA BO DEM CAP NHAT VA THU MUC CAI DAT LOI ($WINDOWS.~BT)
echo ============================================================================
call :DeleteFolder "C:\$WINDOWS.~BT"
call :DeleteFolder "C:\$WINDOWS.~WS"

echo Dang lam moi thu muc SoftwareDistribution va Catroot2...
del /F /Q "C:\ProgramData\Microsoft\Network\Downloader\qmgr*.dat" >nul 2>&1
if exist "C:\Windows\SoftwareDistribution.old" rmdir /S /Q "C:\Windows\SoftwareDistribution.old"
if exist "C:\Windows\System32\catroot2.old" rmdir /S /Q "C:\Windows\System32\catroot2.old"
ren "C:\Windows\SoftwareDistribution" "SoftwareDistribution.old"
ren "C:\Windows\System32\catroot2" "catroot2.old"

echo.
echo ============================================================================
echo   BUOC 3: KHOI DONG LAI DICH VU VA DONG BO MOI TRUONG WINRE (SAFE_OS)
echo ============================================================================
call :StartUpdateServices

echo Dang reset cau hinh Windows Recovery Environment (WinRE)...
reagentc /disable
reagentc /enable
if %errorLevel% neq 0 (
    echo [CANH BAO] reagentc /enable bi loi! Kiem tra lai phan vung Recovery / Winre.wim.
)
reagentc /info

echo.
echo ============================================================================
echo   BUOC 4: KIEM TRA O DIA, DISM VA SFC (MAT KHOANG 10-15 PHUT)
echo ============================================================================
chkdsk C: /scan
DISM /Online /Cleanup-Image /RestoreHealth
sfc /scannow

echo.
echo ============================================================================
echo   HOAN TAT! HAY THAO O SSD CU RA (NEU CON CAM) VA KHOI DONG LAI MAY.
echo ============================================================================
pause
exit /b

:: --- Cac ham dung chung ---
:StopUpdateServices
net stop wuauserv /y >nul 2>&1
net stop cryptSvc /y >nul 2>&1
net stop bits /y >nul 2>&1
net stop msiserver /y >nul 2>&1
net stop UsoSvc /y >nul 2>&1
goto :eof

:StartUpdateServices
net start wuauserv >nul 2>&1
net start cryptSvc >nul 2>&1
net start bits >nul 2>&1
net start msiserver >nul 2>&1
net start UsoSvc >nul 2>&1
goto :eof

:DeleteFolder
if exist "%~1" (
    echo Dang xoa nhanh %~1 ...
    rmdir /S /Q "%~1" >nul 2>&1
)
if exist "%~1" (
    echo   - Con file he thong duoc bao ve, dang cap quyen so huu ^(co the mat vai phut^)...
    takeown /F "%~1" /A /R /D Y >nul 2>&1
    echo   - Dang cap quyen truy cap...
    icacls "%~1" /grant *S-1-5-32-544:F /T /C /Q >nul 2>&1
    echo   - Dang go thuoc tinh chi doc / he thong...
    attrib -R -S -H "%~1\*" /S /D >nul 2>&1
    echo   - Dang xoa not %~1 ...
    rmdir /S /Q "%~1"
)
goto :eof

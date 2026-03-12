@echo off
set VS90COMNTOOLS=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\Tools\
call "%VS90COMNTOOLS%vsvars32.bat"

echo Setting up WinCE ARM build environment...
set PATH=C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\bin\x86_arm;%PATH%
set INCLUDE=C:\Program Files (x86)\Windows CE Tools\SDKs\Toradex_CE700\Include\Armv4i;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\include;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\atlmfc\include;%INCLUDE%
set LIB=C:\Program Files (x86)\Windows CE Tools\SDKs\Toradex_CE700\Lib\ARMV4I;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\lib\armv4i;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\atlmfc\lib\armv4i;%LIB%

echo.
echo =============================================
echo   RTL8152B Driver Build
echo   Target: USB\VID_0BDA^&PID_8152
echo =============================================
echo.

echo Compiling rtl8152.c...
cl.exe /nologo /W3 /O2 ^
    /D "_WIN32_WCE=0x700" ^
    /D "UNDER_CE" ^
    /D "WINCE" ^
    /D "ARM" ^
    /D "_ARM_" ^
    /D "ARMV4I" ^
    /D "UNICODE" ^
    /D "_UNICODE" ^
    /D "NDIS_MINIPORT_DRIVER" ^
    /D "NDIS51_MINIPORT" ^
    /D "_CRT_SECURE_NO_DEPRECATE" ^
    /Fo"rtl8152.obj" /c rtl8152.c
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo === COMPILATION FAILED ===
    exit /b 1
)

echo Linking rtl8152.dll...
link.exe /nologo /DLL ^
    /subsystem:windowsce,7.00 ^
    /machine:THUMB ^
    /NXCOMPAT:NO ^
    /LARGEADDRESSAWARE ^
    /STACK:0x10000,0x1000 ^
    /MERGE:.rdata=.text ^
    /def:rtl8152.def ^
    rtl8152.obj ^
    coredll.lib ^
    corelibc.lib ^
    /out:rtl8152.dll
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo === LINKING FAILED ===
    exit /b 1
)

echo.
echo Checking optional post-link patch step...
set PATCHER=panasonic_pe_patcher.py
if exist "%PATCHER%" (
    echo Applying post-link patch using %PATCHER%...
    python "%PATCHER%" rtl8152.dll
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo === POST-LINK PATCH FAILED ===
        exit /b 1
    )
) else (
    echo No local post-link patcher found. Skipping optional patch step.
)

echo.
echo =============================================
echo   BUILD SUCCESS
echo =============================================
dir rtl8152.dll
echo.
echo Deploy: copy rtl8152.dll to device \Windows\
echo Registry: configure your USB client loader to use rtl8152.dll

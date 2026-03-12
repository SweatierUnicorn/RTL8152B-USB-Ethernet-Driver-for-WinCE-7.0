@echo off
set VS90COMNTOOLS=C:\Program Files (x86)\Microsoft Visual Studio 9.0\Common7\Tools\
call "%VS90COMNTOOLS%vsvars32.bat"

echo Setting up WinCE ARM build environment...
set PATH=C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\bin\x86_arm;%PATH%
set INCLUDE=C:\Program Files (x86)\Windows CE Tools\SDKs\Toradex_CE700\Include\Armv4i;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\include;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\atlmfc\include;%INCLUDE%
set LIB=C:\Program Files (x86)\Windows CE Tools\SDKs\Toradex_CE700\Lib\ARMV4I;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\lib\armv4i;C:\Program Files (x86)\Microsoft Visual Studio 9.0\VC\ce\atlmfc\lib\armv4i;%LIB%

echo.
echo =============================================
echo   NetConfig Build
echo   WinCE 7.0 ARMv4I / THUMB
echo =============================================
echo.

echo Compiling NetConfig.c...
cl.exe /nologo /W3 /O2 ^
    /D "_WIN32_WCE=0x700" ^
    /D "UNDER_CE" ^
    /D "WINCE" ^
    /D "ARM" ^
    /D "_ARM_" ^
    /D "ARMV4I" ^
    /D "UNICODE" ^
    /D "_UNICODE" ^
    /D "_CRT_SECURE_NO_DEPRECATE" ^
    /Fo"NetConfig.obj" /c NetConfig.c
if errorlevel 1 goto compile_fail

echo Linking...
link.exe /nologo ^
    /subsystem:windowsce,7.00 ^
    /machine:THUMB ^
    /entry:WinMainCRTStartup ^
    NetConfig.obj ^
    coredll.lib ^
    corelibc.lib ^
    ws2.lib ^
    iphlpapi.lib ^
    /out:NetConfig.exe
if errorlevel 1 goto link_fail

echo.
echo =============================================
if exist NetConfig.exe (
    echo   SUCCESS: NetConfig.exe created
    for %%F in (NetConfig.exe) do echo   Size: %%~zF bytes
) else (
    echo   WARNING: output file not found
)
echo =============================================
goto done

:compile_fail
echo COMPILE FAILED
goto done

:link_fail
echo LINK FAILED
goto done

:done
echo Done.

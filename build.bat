@echo off
chcp 65001 >nul
echo ================================================================================
echo   PayPal Pagos - PyInstaller Build Script
echo   Build Configuration for Windows Executable
echo ================================================================================
echo.

REM ============================================================================
REM CONFIGURATION
REM ============================================================================

set SCRIPT_DIR=%~dp0
set PROJECT_ROOT=%SCRIPT_DIR%
set PYTHON_PATH=python
set VENV_PATH=%PROJECT_ROOT%\venv

REM Check for virtual environment
if exist "%VENV_PATH%\Scripts\python.exe" (
    set PYTHON_PATH=%VENV_PATH%\Scripts\python.exe
    echo [INFO] Using virtual environment Python
) else (
    echo [INFO] Using system Python (no virtual environment detected)
)

REM ============================================================================
REM STEP 1: INSTALL DEPENDENCIES
REM ============================================================================

echo.
echo ================================================================================
echo STEP 1: Installing Dependencies
echo ================================================================================
echo.

echo Installing required packages from requirements.txt...
%PYTHON_PATH% -m pip install --upgrade pip --quiet
%PYTHON_PATH% -m pip install -r "%PROJECT_ROOT%\requirements.txt" --quiet

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [OK] Dependencies installed successfully
echo.

REM ============================================================================
REM STEP 2: CLEAN PREVIOUS BUILD
REM ============================================================================

echo ================================================================================
echo STEP 2: Cleaning Previous Build
echo ================================================================================
echo.

if exist "%PROJECT_ROOT%\dist" (
    echo Removing old dist folder...
    rmdir /s /q "%PROJECT_ROOT%\dist" >nul 2>&1
    echo [OK] Old dist folder removed
)

if exist "%PROJECT_ROOT%\build" (
    echo Removing old build folder...
    rmdir /s /q "%PROJECT_ROOT%\build" >nul 2>&1
    echo [OK] Old build folder removed
)

echo.

REM ============================================================================
REM STEP 3: RUN PYINSTALLER
REM ============================================================================

echo ================================================================================
echo STEP 3: Building Executable with PyInstaller
echo ================================================================================
echo.

echo Running PyInstaller...
echo.
echo Command: %PYTHON_PATH% -m PyInstaller --windowed --name "PayPalPagos" --specpath "%PROJECT_ROOT%" --distpath "%PROJECT_ROOT%\dist" --workpath "%PROJECT_ROOT%\build" --clean --noupx "%PROJECT_ROOT%\build.spec"
echo.

%PYTHON_PATH% -m PyInstaller ^
    --windowed ^
    --name "PayPalPagos" ^
    --specpath "%PROJECT_ROOT%" ^
    --distpath "%PROJECT_ROOT%\dist" ^
    --workpath "%PROJECT_ROOT%\build" ^
    --clean ^
    --noupx ^
    --log-level WARN ^
    "%PROJECT_ROOT%\build.spec"

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] PyInstaller build failed
    pause
    exit /b 1
)

echo.
echo [OK] PyInstaller build completed successfully
echo.

REM ============================================================================
REM STEP 4: VERIFY BUILD
REM ============================================================================

echo ================================================================================
echo STEP 4: Verifying Build
echo ================================================================================
echo.

set EXE_PATH=%PROJECT_ROOT%\dist\PayPalPagos.exe
if exist "%EXE_PATH%" (
    echo [OK] Executable created successfully!
    echo.
    echo Build location: %EXE_PATH%
    echo.
    echo You can now distribute the PayPalPagos.exe file.
    echo.
) else (
    echo [WARNING] Executable not found at expected location
    echo Checking dist folder contents...
    if exist "%PROJECT_ROOT%\dist" (
        dir "%PROJECT_ROOT%\dist"
    )
)

REM ============================================================================
REM STEP 5: CREATE VERSION INFO FILE (OPTIONAL)
REM ============================================================================

echo.
echo ================================================================================
echo STEP 5: Creating Version Info (Optional)
echo ================================================================================
echo.

set VERSION_INFO="%PROJECT_ROOT%\version_info.txt"
if exist "%VERSION_INFO%" (
    echo [OK] Version info file found
) else (
    echo Creating default version_info.txt...
    (
        echo # Version Information for PayPalPagos
        echo VSVersionInfo(
        echo   ffi=FixedFileInfo(
        echo     filevers=(1, 0, 0, 0),
        echo     prodvers=(1, 0, 0, 0),
        echo     mask=0x3f,
        echo     flags=0x0,
        echo     OS=0x40004,
        echo     type=0x1,
        echo     subsys=0x0,
        echo     language=0x0,
        echo     signature=0xFEEF04BD,
        echo   ),
        echo   kids=[
        echo     StringFileInfo(
        echo       [
        echo         StringTable(
        echo           0x0409, 0x04E4,
        echo           [
        echo             StringStruct(u'CompanyName', u'Your Company'),
        echo             StringStruct(u'FileDescription', u'PayPal Pagos Automation'),
        echo             StringStruct(u'FileVersion', u'1.0.0'),
        echo             StringStruct(u'InternalName', u'PayPalPagos'),
        echo             StringStruct(u'LegalCopyright', u'Copyright © 2024'),
        echo             StringStruct(u'OriginalFilename', u'PayPalPagos.exe'),
        echo             StringStruct(u'ProductName', u'PayPal Pagos'),
        echo             StringStruct(u'ProductVersion', u'1.0.0'),
        echo           ]
        echo         )
        echo       ]
        echo     ),
        echo     VarFileInfo([VarStruct(u'Translation', [0x0409, 0x04E4])])
        echo   ]
        echo )
    ) > "%VERSION_INFO%"
    echo [OK] Default version_info.txt created
)

REM ============================================================================
REM SUMMARY
REM ============================================================================

echo.
echo ================================================================================
echo BUILD COMPLETE
echo ================================================================================
echo.
echo Summary:
echo   - Executable: %PROJECT_ROOT%\dist\PayPalPagos.exe
echo   - Build folder: %PROJECT_ROOT%\build
echo   - Distribution folder: %PROJECT_ROOT%\dist
echo.
echo To rebuild:
echo   1. Run this script again
echo   2. Or run: %PYTHON_PATH% -m PyInstaller build.spec
echo.
echo Note: The .exe is portable and can be distributed without Python installation.
echo       However, it requires the Visual C++ Redistributable to run.
echo.
pause
exit /b 0

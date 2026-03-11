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
set VENV_PATH=%PROJECT_ROOT%venv
set OUTPUT_DIR=O:\Finanzas\Info Bancos\Pagos Internacionales\PAYPAL

REM ============================================================================
REM STEP 1: INSTALL DEPENDENCIES
REM ============================================================================

echo.
echo ================================================================================
echo STEP 1: Installing Dependencies
echo ================================================================================
echo.

echo Installing required packages from requirements.txt...
"%PYTHON_PATH%" -m pip install --upgrade pip --quiet
"%PYTHON_PATH%" -m pip install -r "%PROJECT_ROOT%requirements.txt" --quiet

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [OK] Dependencies installed successfully
echo.

echo ================================================================================
echo STEP 2: Cleaning Previous Build
echo ================================================================================
echo.

if exist "%PROJECT_ROOT%build" (
    echo Removing old build folder...
    rmdir /s /q "%PROJECT_ROOT%build" >nul 2>&1
    if exist "%PROJECT_ROOT%build" (
        echo [WARNING] Could not remove build folder completely
    ) else (
        echo [OK] Old build folder removed
    )
)

REM Remove old spec file if exists
if exist "%PROJECT_ROOT%pagosPaypal.spec" (
    echo Removing old spec file...
    del /f /q "%PROJECT_ROOT%pagosPaypal.spec" >nul 2>&1
    echo [OK] Old spec file removed
)

echo.

echo ================================================================================
echo STEP 3: Verifying Required Files
echo ================================================================================
echo.

REM Check for main Python files
if not exist "%PROJECT_ROOT%interfaz.py" (
    echo [ERROR] interfaz.py not found!
    pause
    exit /b 1
)
echo [OK] interfaz.py found

if not exist "%PROJECT_ROOT%main.py" (
    echo [ERROR] main.py not found!
    pause
    exit /b 1
)
echo [OK] main.py found
if not exist "%PROJECT_ROOT%scripts\verificacion.py" (
    echo [ERROR] verificacion.py not found!
    pause
    exit /b 1
)
echo [OK] verificacion.py found

REM Check for optional logo.ico
if exist "%PROJECT_ROOT%logo.ico" (
    echo [OK] logo.ico found
    set HAS_ICON=1
) else (
    echo [WARNING] logo.ico not found - executable will not have a custom icon
    set HAS_ICON=0
)

echo.

REM ============================================================================
REM STEP 4: RUN PYINSTALLER
REM ============================================================================

echo ================================================================================
echo STEP 4: Building Executable with PyInstaller
echo ================================================================================
echo.

echo Running PyInstaller...
echo.

REM Build command with conditional icon
if "%HAS_ICON%"=="1" (
    echo [INFO] Building with custom icon...
    "%PYTHON_PATH%" -m PyInstaller ^
        --icon=logo.ico ^
        --add-data "logo.ico;." ^
        --windowed ^
        --onedir ^
        --name "Pagos paypal RPA" ^
        --distpath "%OUTPUT_DIR%" ^
        --workpath "%PROJECT_ROOT%build" ^
        --clean ^
        --noupx ^
        --log-level WARN ^
        --hidden-import=selenium ^
        --hidden-import=pandas ^
        --hidden-import=openpyxl ^
        --hidden-import=fitz ^
        --hidden-import=customtkinter ^
        --collect-all customtkinter ^
        interfaz.py
) else (
    echo [INFO] Building without custom icon...
    "%PYTHON_PATH%" -m PyInstaller ^
        --windowed ^
        --onedir ^
        --name "Pagos paypal RPA" ^
        --distpath "%OUTPUT_DIR%" ^
        --workpath "%PROJECT_ROOT%build" ^
        --clean ^
        --noupx ^
        --log-level WARN ^
        --hidden-import=selenium ^
        --hidden-import=pandas ^
        --hidden-import=openpyxl ^
        --hidden-import=fitz ^
        --hidden-import=customtkinter ^
        --collect-all customtkinter ^
        interfaz.py
)

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] PyInstaller build failed
    echo Check the errors above for more details
    pause
    exit /b 1
)

echo.
echo [OK] PyInstaller build completed successfully
echo.

REM ============================================================================
REM STEP 5: VERIFY BUILD
REM ============================================================================

echo ================================================================================
echo STEP 5: Verifying Build
echo ================================================================================
echo.

set EXE_PATH=%OUTPUT_DIR%\Pagos paypal RPA\Pagos paypal RPA.exe
if exist "%EXE_PATH%" (
    echo [OK] Executable created successfully!
    echo.
    echo Build location: %EXE_PATH%
    echo.
    
    REM Get file size
    for %%A in ("%EXE_PATH%") do set SIZE=%%~zA
    echo File size: %SIZE% bytes
    echo.
    
    echo You can now distribute the Pagos paypal RPA folder with all its contents.
    echo.
) else (
    echo [WARNING] Executable not found at expected location: %EXE_PATH%
    echo.
    echo Checking dist folder contents...
    if exist "%OUTPUT_DIR%" (
        dir "%OUTPUT_DIR%" /s
    ) else (
        echo [ERROR] Output directory does not exist!
    )
)


echo.
echo ================================================================================
echo BUILD COMPLETE
echo ================================================================================
echo.
echo Summary:
echo   - Executable: %EXE_PATH%
echo   - Build folder: %PROJECT_ROOT%build
echo   - Distribution folder: %OUTPUT_DIR%
echo   - README: %README_PATH%
echo   - Default output: O:\Finanzas\Info Bancos\Pagos Internacionales\PAYPAL
echo.
echo To rebuild:
echo   1. Run this script again
echo   2. Or run: "%PYTHON_PATH%" -m PyInstaller Pagos paypal RPA.spec
echo.
echo IMPORTANT NOTES:
echo   - The executable is portable and includes all dependencies
echo   - Distribute the entire Pagos paypal RPA folder, not just the .exe
echo   - Users may need Visual C++ Redistributable installed
echo   - The first run may be slower as it unpacks dependencies
echo.
echo To test the executable:
echo   1. Navigate to: %OUTPUT_DIR%\Pagos paypal RPA
echo   2. Run: Pagos paypal RPA.exe
echo.
echo NOTE: By default, the executable is created in:
echo       O:\Finanzas\Info Bancos\Pagos Internacionales\PAYPAL\Pagos paypal RPA\
echo.
pause
exit /b 0

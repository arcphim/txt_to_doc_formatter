@echo off
chcp 65001 >nul

REM Check Python installation
echo Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python not found. Please install Python 3.7+ and add to PATH.
    pause
    exit /b 1
)

echo Python is ready.

REM Check pip installation
echo Checking pip...
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: pip not found. Please install pip and add to PATH.
    pause
    exit /b 1
)

echo pip is ready.

REM Check PyInstaller
echo Checking PyInstaller...
python -c "import PyInstaller" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing PyInstaller...
    pip install PyInstaller
    if %errorlevel% neq 0 (
        echo Error: Failed to install PyInstaller. Check network connection.
        pause
        exit /b 1
    )
) else (
    echo PyInstaller is ready.
)

REM Clean old files
echo Cleaning old files...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

REM Check config.json
if not exist "config.json" (
    echo Error: config.json not found. Please create config.json before building.
    pause
    exit /b 1
)

echo config.json is ready.

REM Build application
echo Building application...
python -m PyInstaller --noconfirm --onefile --windowed --icon="app.ico" --name="文件转公文" "main.py"
if %errorlevel% neq 0 (
    echo Error: Failed to build application.
    pause
    exit /b 1
)

REM Copy files
if exist "dist\文件转公文.exe" (
    echo Copying necessary files...
    copy /Y "config.json" "dist\" >nul 2>&1
    echo Build completed! Application is in dist directory as "文件转公文.exe" with "config.json"
) else (
    echo Error: Output directory not found.
    pause
    exit /b 1
)

pause
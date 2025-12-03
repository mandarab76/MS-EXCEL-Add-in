@echo off
REM ==============================
REM ATS Financials Add-in Installer
REM ==============================

REM Step 1: Change to add-in directory (where this batch file is located)
cd /d "%~dp0"

REM Step 2: Check if package.json exists
if not exist "package.json" (
    echo ERROR: package.json not found!
    echo Please ensure this batch file is in your Office.js add-in root directory.
    pause
    exit /b
)

REM Step 3: Check if manifest.xml exists
if not exist "manifest.xml" (
    echo ERROR: manifest.xml not found!
    echo Please ensure the manifest file is present for Excel sideloading.
    pause
    exit /b
)

REM Step 4: Install npm dependencies
echo Installing required npm packages from package.json...
npm install

REM Step 5: Start the Office.js dev server
echo Starting local Office.js development server...
start cmd /k "npm start"

REM Step 6: Instructions for Excel sideloading
echo ================================
echo Setup Complete!
echo 1. Open Excel Desktop.
echo 2. Go to Insert > Get Add-ins > Shared Folder.
echo 3. Select manifest.xml from this directory.
echo 4. Use 'Financial Statements' tab in Excel.
echo ================================
pause
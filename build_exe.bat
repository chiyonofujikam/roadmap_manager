@echo off
REM Build roadmap.exe using PyInstaller
REM Make sure PyInstaller is installed: pip install pyinstaller
REM Or use: uv pip install pyinstaller

echo Building roadmap.exe...
echo.

pyinstaller --onefile ^
    --name roadmap ^
    --console ^
    --clean ^
    --noupx ^
    --optimize=2 ^
    --paths . ^
    --hidden-import=roadmap ^
    --hidden-import=roadmap.helpers ^
    --hidden-import=roadmap.main ^
    --hidden-import=roadmap.roadmap ^
    --hidden-import=openpyxl ^
    --hidden-import=tqdm ^
    roadmap_cli.py

echo.
if exist "dist\roadmap.exe" (
    echo SUCCESS! Executable created at: dist\roadmap.exe
    echo.

    REM Create destination directory if it doesn't exist
    if not exist "C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script" (
        mkdir "C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script"
        echo Created destination directory: C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\files\script
    )

    REM Check if file already exists at destination
    if exist "C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script\roadmap.exe" (
        echo Existing roadmap.exe found at destination. It will be overwritten.
    )

    REM Copy the executable to destination (/Y flag overwrites without prompting)
    copy /Y "dist\roadmap.exe" "C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script\roadmap.exe"

    if exist "C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script\roadmap.exe" (
        echo.
        echo Executable copied to: C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\files\script\roadmap.exe
    ) else (
        echo WARNING: Failed to copy executable to C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\files\script
    )

    echo.
    echo Build complete! Executable is available at:
    echo   - dist\roadmap.exe (original)
    echo   - C:\Users\MustaphaELKAMILI\OneDrive - IKOSCONSULTING\test_RM\roadmap_manager\script\roadmap.exe (copied)
) else (
    echo ERROR: Build failed. Check the output above for errors.
)
echo.
pause

@echo off
echo ========================================
echo BASH Flow VBA Module Setup
echo ========================================
echo.
echo This script helps set up your VBA modules for Outlook.
echo.
echo Created .bas files for direct import:
echo.

if exist "BASH_Flow_Sandbox.bas" (
    echo [OK] BASH_Flow_Sandbox.bas
) else (
    echo [MISSING] BASH_Flow_Sandbox.bas
)

if exist "WorkflowManager.bas" (
    echo [OK] WorkflowManager.bas
) else (
    echo [MISSING] WorkflowManager.bas
)

if exist "WorkflowDashboard.bas" (
    echo [OK] WorkflowDashboard.bas
) else (
    echo [MISSING] WorkflowDashboard.bas
)

if exist "DashboardSetup.bas" (
    echo [OK] DashboardSetup.bas
) else (
    echo [MISSING] DashboardSetup.bas
)

if exist "SandboxDiagnostic.bas" (
    echo [OK] SandboxDiagnostic.bas
) else (
    echo [MISSING] SandboxDiagnostic.bas
)

if exist "CheckDatabase.bas" (
    echo [OK] CheckDatabase.bas
) else (
    echo [MISSING] CheckDatabase.bas
)

if exist "QuickSetup.bas" (
    echo [OK] QuickSetup.bas
) else (
    echo [MISSING] QuickSetup.bas
)

if exist "RibbonSetup.bas" (
    echo [OK] RibbonSetup.bas
) else (
    echo [MISSING] RibbonSetup.bas
)

if exist "BASH_Flow_Ribbon.exportedUI" (
    echo [OK] BASH_Flow_Ribbon.exportedUI
) else (
    echo [MISSING] BASH_Flow_Ribbon.exportedUI
)

echo.
echo ========================================
echo Setup Instructions:
echo ========================================
echo.
echo 1. Open Outlook
echo 2. Press Alt+F11 to open VBA Editor
echo 3. Right-click on your project
echo 4. Select "Import File..."
echo 5. Import each .bas file from this folder
echo.
echo OR run QuickSetup.bas in VBA Editor:
echo - Import QuickSetup.bas first
echo - Then run: SetupBASHFlowModules()
echo.
echo ========================================
echo Files ready for import!
echo ========================================
echo.
pause

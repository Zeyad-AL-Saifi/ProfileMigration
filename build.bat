@echo off
cd /d "C:\Users\Suhayb\Desktop\REPOS\ProfileMigration"
dotnet build ProfileMigration/ProfileMigration.App.csproj -v minimal 2>&1 > build_output.txt
echo. >> build_output.txt
echo Exit code: %ERRORLEVEL% >> build_output.txt
echo Build complete. Output saved to build_output.txt
pause

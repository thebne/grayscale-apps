@echo off
echo Creating startup shortcut for Grayscale Apps...

powershell -Command ^
  "$ws = New-Object -ComObject WScript.Shell;" ^
  "$startup = $ws.SpecialFolders('Startup');" ^
  "$shortcut = $ws.CreateShortcut(\"$startup\Grayscale Apps.lnk\");" ^
  "$shortcut.TargetPath = '%~dp0start_background.bat';" ^
  "$shortcut.WorkingDirectory = '%~dp0';" ^
  "$shortcut.WindowStyle = 7;" ^
  "$shortcut.Description = 'Per-window grayscale overlay';" ^
  "$shortcut.Save();" ^
  "Write-Host 'Shortcut created at:' \"$startup\Grayscale Apps.lnk\""

echo Done. Grayscale Apps will now start on login.
pause

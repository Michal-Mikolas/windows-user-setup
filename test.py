import time, subprocess
from windows.windows import Windows
import config
import time

windows = Windows(config.cache_path)

windows.run_powershell_command('$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Chrome.lnk"); $Shortcut.TargetPath = "C:\Program Files\Google\Chrome\Application\chrome.exe"; $Shortcut.Save()')

windows.run_powershell_command('$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Outlook.lnk"); $Shortcut.TargetPath = "C:\Program Files\Microsoft Office\Office14\OUTLOOK.EXE"; $Shortcut.Save()')

windows.run_powershell_command('$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Word.lnk"); $Shortcut.TargetPath = "C:\Program Files\Microsoft Office\Office14\WINWORD.EXE"; $Shortcut.Save()')

windows.run_powershell_command('$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut("$Home\Desktop\Excel.lnk"); $Shortcut.TargetPath = "C:\Program Files\Microsoft Office\Office14\EXCEL.EXE"; $Shortcut.Save()')

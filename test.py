import time, subprocess
from windows.windows import Windows
import config
import time

windows = Windows(config.cache_path)

print('TEST start')
subprocess.call([r'%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe', r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force'], shell=True)
# subprocess.run(['explorer.exe', 'ms-settings:regionlanguage'])
# subprocess.Popen(['explorer.exe', 'ms-settings:regionlanguage'])
print('TEST end')
time.sleep(3)

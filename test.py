import time, subprocess, winreg
from winregistry import WinRegistry
from windows.windows import Windows
import config

windows = Windows(config.cache_path)

windows.uninstall('Java')

reg = windows.load_registry(r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
print(reg)

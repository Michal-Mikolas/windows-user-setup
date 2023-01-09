import time, subprocess
from windows.windows import Windows
import config
import time

windows = Windows(config.cache_path)

windows.copy(
	'\\\\10.0.0.12\\all\\_INSTALL\\Data\\ChromeProfile\\**',
	'%LOCALAPPDATA%\\Google\\Chrome TEST\\User Data'
)

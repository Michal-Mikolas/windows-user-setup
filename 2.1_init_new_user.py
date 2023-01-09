import time
from windows.windows import Windows
from windows.settings import Settings
import config
from datetime import datetime

def log(msg):
	print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

windows = Windows(config.cache_path)


#     #
#  #  #  ####  #####  #    #
#  #  # #    # #    # #   #
#  #  # #    # #    # ####
#  #  # #    # #####  #  #
#  #  # #    # #   #  #   #
 ## ##   ####  #    # #    #
log('Unifying Windows UI')

log('- setting up tray icons to be all visible')
windows.unhide_tray_icons()

log('- showing file extensions in file explorer')
windows.unhide_file_extensions()

log('- showing This PC on desktop')
windows.show_this_pc()

log('- disabling Cortana')
windows.disable_cortana()

log('- enabling dark theme')
windows.enable_dark_theme()

log('- disabling Windows transparency')
windows.disable_transparency()

log('- installing languages: en-US, cs-CZ')
windows.setup_languages('en-US,cs-CZ')

log('- setting display language to en-US')
windows.set_display_language('en-US')

log('- adding shortcuts to Desktop')
# TODO

log('- clearing Desktop')
windows.clear_desktop(keep=['acrobat', 'chrome', 'backup', 'tor', 'nn.xlsx', 'ED7BA470-8E54-465E-825C-99712043E01C', 'vp2', 'moba'])

log('- setting up Chrome browser')
chrome_home = '%LOCALAPPDATA%\\Google\\Chrome\\User Data'
if not windows.exists(f'{chrome_home}\\Default') and not windows.exists(f'{chrome_home}\\First Run'):
	windows.copy('\\\\10.0.0.12\\all\\_INSTALL\\Data\\ChromeProfile\\**', chrome_home)

log('- done')

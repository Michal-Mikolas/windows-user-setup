import time, glob, os
from windows.windows import Windows
from windows.settings import Settings
import config
from datetime import datetime

def log(msg):
	print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

windows = Windows(config.cache_path)
windows.chdir('%UserProfile%')


#     #
#  #  #  ####  #####  #    #
#  #  # #    # #    # #   #
#  #  # #    # #    # ####
#  #  # #    # #####  #  #
#  #  # #    # #   #  #   #
 ## ##   ####  #    # #    #
log('Preparing Windows user environment')

#
log('- setting up tray icons to be all visible')
windows.unhide_tray_icons()

#
log('- showing file extensions in file explorer')
windows.unhide_file_extensions()

#
log('- showing This PC on desktop')
windows.show_this_pc()

#
log('- disabling Cortana')
windows.disable_cortana()

#
log('- enabling dark theme')
windows.enable_dark_theme()

#
log('- disabling Windows transparency')
windows.disable_transparency()

#
log('- installing languages: en-US, cs-CZ')
windows.setup_languages('en-US,cs-CZ')

#
log('- setting display language to en-US')
windows.set_display_language('en-US')

#
log('- adding shortcuts to Desktop')
apps = {
	'Adobe Acrobat': [
		'C:\\Program Files*\\Adobe\\Acrobat*\\Reader\\AcroRd*.exe',
		'C:\\Program Files*\\Adobe\\Acrobat*\\Acrobat\\Acrobat.exe',
	],
	'Chrome': [
		'C:\\Program Files*\\Google\\Chrome\\Application\\chrome.exe',
	],
	'Outlook': [
		'C:\\Program Files*\\Microsoft Office\\Office*\\OUTLOOK.EXE',
		'C:\\Program Files*\\Microsoft Office\\root\\Office*\\OUTLOOK.EXE',
	],
	'Word': [
		'C:\\Program Files*\\Microsoft Office\\Office*\\WINWORD.EXE',
		'C:\\Program Files*\\Microsoft Office\\root\\Office*\\WINWORD.EXE',
	],
	'Excel': [
		'C:\\Program Files*\\Microsoft Office\\Office*\\EXCEL.EXE',
		'C:\\Program Files*\\Microsoft Office\\root\\Office*\\EXCEL.EXE',
	],
}
for (name, paths) in apps.items():
	print(f'- - {name}')

	if windows.has_desktop_shortcut(name):
		print(f'- - - already exists')
		continue

	file = None
	for path in paths:
		for file in glob.glob(path, recursive=True):
			print(f'- - - adding shortcut to "{file}"')
			windows.create_desktoop_shortcut(file, name)
			break

		if file:
			break

	if not file:
		print(f'- - - app not found :-(')

#
log('- clearing Desktop')
windows.clear_desktop(keep=['Adobe Acrobat', 'Chrome', 'Outlook', 'Word', 'Excel', 'backup', 'tor', 'nn.xlsx', 'ED7BA470-8E54-465E-825C-99712043E01C', 'vp2', 'moba'])

#
log('- setting up Chrome browser')
chrome_home = '%LOCALAPPDATA%\\Google\\Chrome\\User Data'
if not windows.exists(f'{chrome_home}\\Default') and not windows.exists(f'{chrome_home}\\First Run'):
	windows.copy('\\\\10.0.0.12\\all\\_INSTALL\\Data\\ChromeProfile\\**', chrome_home)
	log('- - done')
else:
	log('- - nothing to do, Chrome was already run')

#
log('- uninstalling apps')
apps = ['Seznam', 'WinRar', 'WinZip', 'Dropbox', 'OneDrive', 'McAfee', 'Backup and Sync']
for app in apps:
	log(f'- - {app}')
	windows.uninstall(app, logger=lambda msg: log(f'- - - {msg}'))

#
log('- done')

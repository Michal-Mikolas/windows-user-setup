import glob
from windows.windows import Windows
import config
from datetime import datetime

def log(msg):
	print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

try:
	# windows = Windows(config.cache_path)
	windows = Windows()
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

	# #
	# log('- enabling dark theme')
	# windows.enable_dark_theme()

	# #
	# log('- disabling Windows transparency')
	# windows.disable_transparency()

	#
	log('- installing languages: en-US, cs-CZ')
	windows.setup_languages('en-US,cs-CZ')

	#
	log('- setting display language to en-US')
	windows.set_display_language('en-US')

	#
	log('- clearing Desktop')
	windows.clear_desktop(keep=['Edge', 'Outlook', 'Word', 'Excel', 'backup', 'tor', 'nn.xlsx', 'ED7BA470-8E54-465E-825C-99712043E01C', 'vp2', 'moba', 'Mattermost'])

	#
	log('- adding shortcuts to Desktop')
	apps = {
		# 'Chrome': [
		# 	'C:\\Program Files*\\Google\\Chrome\\Application\\chrome.exe',
		# ],
		# 'Adobe Acrobat': [
		# 	'C:\\Program Files*\\Adobe\\Acrobat*\\Reader\\AcroRd*.exe',
		# 	'C:\\Program Files*\\Adobe\\Acrobat*\\Acrobat\\Acrobat.exe',
		# ],
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
		'Mattermost': [
			f'C:\\Users\\{os.getlogin()}\\AppData\\Local\\Programs\\mattermost-desktop\\Mattermost.exe',
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
	log('- running apps for setup')
	apps = {
		'Edge': [
			'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
		],
		'Excel': [
			'C:\\Program Files*\\Microsoft Office\\Office*\\EXCEL.EXE',
			'C:\\Program Files*\\Microsoft Office\\root\\Office*\\EXCEL.EXE',
		],
		'Outlook': [
			'C:\\Program Files*\\Microsoft Office\\Office*\\OUTLOOK.EXE',
			'C:\\Program Files*\\Microsoft Office\\root\\Office*\\OUTLOOK.EXE',
		],
		'Mattermost': [
			f'C:\\Users\\{os.getlogin()}\\AppData\\Local\\Programs\\mattermost-desktop\\Mattermost.exe',
		],
	}
	for (name, paths) in apps.items():
		print(f'- - {name}')

		file = None
		for path in paths:
			for file in glob.glob(path, recursive=True):
				print(f'- - - running "{file}"')
				windows.run(file)

				if 'Edge' in file:
					print('')
					input('Press ENTER to continue...')
					print('')

				break

			if file:
				break

		if not file:
			print(f'- - - NOT FOUND')

	# #
	# log('- setting up Chrome browser')
	# chrome_home = '%LOCALAPPDATA%\\Google\\Chrome\\User Data'
	# if not windows.exists(f'{chrome_home}\\Default') and not windows.exists(f'{chrome_home}\\First Run'):
	# 	windows.copy('\\\\10.0.0.12\\all\\_INSTALL\\Data\\ChromeProfile\\**', chrome_home)
	# 	log('- - done')
	# else:
	# 	log('- - nothing to do, Chrome was already run')

	#
	log('- setting up printer')
	# apps = ['Seznam', 'WinRar', 'WinZip', 'Dropbox', 'OneDrive', 'McAfee', 'Backup and Sync']
	# for app in apps:
	# 	log(f'- - {app}')
	# 	windows.uninstall(app, logger=lambda msg: log(f'- - - {msg}'))
	windows.settings_printers()

	print('')
	input('Press ENTER to continue...')
	print('')

	#
	log('- uninstalling apps')
	# apps = ['Seznam', 'WinRar', 'WinZip', 'Dropbox', 'OneDrive', 'McAfee', 'Backup and Sync']
	# for app in apps:
	# 	log(f'- - {app}')
	# 	windows.uninstall(app, logger=lambda msg: log(f'- - - {msg}'))
	windows.settings_apps()

	print('')
	input('Press ENTER to continue...')
	print('')

	#
	log('- done')

	# Next steps tips
	print('')
	log('- Tips: ')
	log('- - Switch Windows language to Czech')
	log('- - Manage startup apps')
	log('- - Restart PC')

except Exception as e:
	log(f'{type(e).__name__}: {str(e)}')

print('')
input('Press ENTER to continue...')

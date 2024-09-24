import os, sys, re, platform, subprocess, glob, shutil, winreg

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


"""
Sources:
https://winaero.com/ms-settings-commands-in-windows-10/
http://woshub.com/ms-settings-uri-commands-windows-11/
https://support.microsoft.com/en-us/windows/windows-keyboard-shortcuts-3d444b08-3a00-abd6-67da-ecfc07e86b98
https://support.microsoft.com/en-us/windows/keyboard-shortcuts-in-windows-dcc61a57-8ff0-cffe-9796-cb9706c75eec#WindowsVersion=Windows_10
https://www.dasm.cz/clanek/jak-z-windows-10-udelat-desktopovy-system
https://www.dasm.cz/clanek/jak-z-windows-10-udelat-desktopovy-system-ii
https://superuser.com/questions/217504/is-there-a-list-of-windows-special-directories-shortcuts-like-temp
https://community.spiceworks.com/how_to/1142-obtain-an-uninstall-string-for-any-application
https://serverfault.com/questions/950291/how-uninstall-a-program-using-the-uninstallstring-found-in-regedit-with-cmd-msie
"""
class Windows():

	def __init__(self, matt_cache_path:str = None) -> None:
		if matt_cache_path:
			from matt.matt import Matt

			self.matt = Matt(cache_file=matt_cache_path)

			dir = os.path.realpath(os.path.dirname(__file__)) + '/img'  # C:/windows-user-setup/windows/img
			dir = dir.replace(os.getcwd(), '').lstrip('\\/')            # windows/img
			self.matt.set_ui({
				'run_title': [
					f'{dir}/run_title_w10.png',
					f'{dir}/run_title_w10.2.png',
					f'{dir}/run_title_w11.png',
					f'{dir}/run_title_w11.2.png',
				],
				'search_app_done': [
					f'{dir}/search_app_done_w10_dark.png',
					f'{dir}/search_app_done_selected_w10_dark.png',
					f'{dir}/search_recommended_done_w10_dark.png',
					f'{dir}/search_recommended_done_selected_w10_dark.png',
					f'{dir}/search_app_done_w11_dark.png',
					f'{dir}/search_app_done_selected_w11_dark.png',
					f'{dir}/search_recommended_done_w11_dark.png',
					f'{dir}/search_recommended_done_selected_w11_dark.png',
				],
				'search_ready': [
					f'{dir}/search_ready_w10_dark.png',
					f'{dir}/search_ready_w11_dark.png',
				],
				'taskbar_option_news_off': [
					f'{dir}/taskbar_option_news_off_w10_dark.png',
				],
				'taskbar_option_news': [
					f'{dir}/taskbar_option_news_w10_dark.png',
				],
				'taskbar_option_search_icon': [
					f'{dir}/taskbar_option_search_icon_w10_dark.png',
				],
				'taskbar_option_search': [
					f'{dir}/taskbar_option_search_w10_dark.png',
				],
				'taskbar': [
					f'{dir}/taskbar_w10_dark.png',
				],
				'unpin_from_taskbar': [
					f'{dir}/unpin_from_taskbar_w10_dark.png',
					f'{dir}/unpin_from_taskbar_w11_dark.png',
				],
			})

		self.winver = None

	def is_win_10(self):
		# https://stackoverflow.com/questions/66554824/get-windows-version-in-python
		version = int( platform.version().split('.')[2] )

		return (19000 <= version <= 21999)


	######                               #####
	#     #  ####  #    # ###### #####  #     # #    # ###### #      #            #####    ##    ####  ###### #####
	#     # #    # #    # #      #    # #       #    # #      #      #            #    #  #  #  #      #      #    #
	######  #    # #    # #####  #    #  #####  ###### #####  #      #            #####  #    #  ####  #####  #    #
	#       #    # # ## # #      #####        # #    # #      #      #            #    # ######      # #      #    #
	#       #    # ##  ## #      #   #  #     # #    # #      #      #            #    # #    # #    # #      #    #
	#        ####  #    # ###### #    #  #####  #    # ###### ###### ######       #####  #    #  ####  ###### #####
	CMD_SHOW_TRAY_ICONS = r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer -Name EnableAutoTray -Value 0 -Type Dword -Force'
	CMD_SHOW_FILE_EXTENSIONS = r'Set-ItemProperty -Path HKCU:\HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced -Name HideFileExt -Value 0 -Type Dword -Force'
	CMD_SHOW_THIS_PC = r'If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu")) { ' \
		+ r'	New-Item -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu" | Out-Null; ' \
		+ r'}; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\ClassicStartMenu -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0 -Type Dword -Force; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel -Name "{20D04FE0-3AEA-1069-A2D8-08002B30309D}" -Value 0 -Type Dword -Force'
	CMD_CORTANA_OFF = r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Personalization\Settings -Name AcceptedPrivacyPolicy -Value 0 -Type Dword -Force; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\InputPersonalization -Name RestrictImplicitTextCollection -Value 1 -Type Dword -Force; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\InputPersonalization -Name RestrictImplicitInkCollection -Value 1 -Type Dword -Force; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore -Name HarvestContacts -Value 0 -Type Dword -Force'
	CMD_DARK_THEME_ON = r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force; ' \
		+ r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force'
	CMD_TRANSPARENCY_OFF = r'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name EnableTransparency -Value 0 -Type Dword -Force'
	CMD_SETUP_LANGUAGES = r'Set-WinUserLanguageList -Force -LanguageList {}'
	CMD_SET_DISPLAY_LANGUAGE = r'Set-WinUILanguageOverride -Language {}'
	CMD_CREATE_DESKTOP_SHORTCUT = r'$WshShell = New-Object -comObject WScript.Shell; ' \
		+ r'$Shortcut = $WshShell.CreateShortcut("$Home\Desktop\{}.lnk"); ' \
		+ r'$Shortcut.TargetPath = "{}"; ' \
		+ r'$Shortcut.Save()'

	def run(self, command:str):
		subprocess.run(command, shell=True)

	def run_powershell_command(self, command):
		subprocess.run([
			r'%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe',
			command
		], shell=True)

	def unhide_tray_icons(self):
		self.run_powershell_command(self.CMD_SHOW_TRAY_ICONS)

	def unhide_file_extensions(self):
		self.run_powershell_command(self.CMD_SHOW_FILE_EXTENSIONS)

	def disable_transparency(self):
		self.run_powershell_command(self.CMD_TRANSPARENCY_OFF)

	def show_this_pc(self):
		self.run_powershell_command(self.CMD_SHOW_THIS_PC)

	def disable_cortana(self):
		self.run_powershell_command(self.CMD_CORTANA_OFF)

	def enable_dark_theme(self):
		self.run_powershell_command(self.CMD_DARK_THEME_ON)

	def setup_languages(self, languages):
		# http://www.lingoes.net/en/translator/langcode.htm
		self.run_powershell_command(self.CMD_SETUP_LANGUAGES.format(languages))

	def set_display_language(self, language):
		# http://www.lingoes.net/en/translator/langcode.htm
		self.run_powershell_command(self.CMD_SET_DISPLAY_LANGUAGE.format(language))

	def create_desktoop_shortcut(self, dest:str, name:str):
		self.run_powershell_command(self.CMD_CREATE_DESKTOP_SHORTCUT.format(name, dest))

	def uninstall(self, mask:str, logger=None):
		found_in_registry = self.find_uninstall_script_in_registry(mask, logger=logger)
		found_in_files = self.find_uninstall_script_in_files(mask)

		if found_in_registry:
			for cmd in found_in_registry:
				logger(f'> {cmd}')
				self.run(cmd)

		elif found_in_files:
			for cmd in found_in_files:
				logger(f'> {cmd}')
				self.run(cmd)

		else:
			logger('no uninstall script found')

	def settings_apps(self):
		subprocess.run(['explorer.exe', 'ms-settings:appsfeatures'])

	def settings_printers(self):
		subprocess.run(['explorer.exe', 'ms-settings:printers'])


	######
	#     # ######  ####  #  ####  ##### #####  #   #
	#     # #      #    # # #        #   #    #  # #
	######  #####  #      #  ####    #   #    #   #
	#   #   #      #  ### #      #   #   #####    #
	#    #  #      #    # # #    #   #   #   #    #
	#     # ######  ####  #  ####    #   #    #   #
	def find_uninstall_script_in_registry(self, mask, logger=None):
		mask = mask.replace('*', '.*')
		logger = logger or (lambda v: v)
		found = []

		paths = [
			r'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
			r'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
		]
		for path in paths:
			reg = self.load_registry(path)

			for name, key in reg['keys'].items():
				display_name = key.get('values', {}).get('DisplayName', '')

				if re.search(mask, name) or re.search(mask, display_name):
					logger(f'{key["path"]} ({display_name})')

					uninstall_string = key.get('values', {}).get('UninstallString', '')
					quiet_uninstall_string = key.get('values', {}).get('QuietUninstallString', '')
					if uninstall_string or quiet_uninstall_string:
						cmd = quiet_uninstall_string or uninstall_string
						cmd = self.fix_uninstall_cmd(cmd)
						found.append(cmd)

		return found

	def load_registry(self, path:str):
		result = {'keys': {}, 'values': {}}

		root_folder = path.split('\\')[0]
		path = re.sub(r'^[^\\]+\\', '', path)
		r_conn = winreg.ConnectRegistry(None, getattr(winreg, root_folder))

		#
		# Keys
		#
		r_key = winreg.OpenKey(r_conn, path)
		i = 0
		while True:
			try:
				name = winreg.EnumKey(r_key, i)

				children = self.load_registry(f'{root_folder}\{path}\{name}')

				result['keys'][name] = {
					'name': name,
					'path': f'{root_folder}\{path}\{name}',
					'keys': children['keys'],
					'values': children['values'],
				}

				# key2 = winreg.OpenKey(reg, rf'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{x}')
				# for i in range(0, 5):
				# 	value = winreg.EnumValue(key2, i)
				# 	print(f'- value: {value}')

				i += 1
			except OSError as e:
				if '[WinError 259]' not in str(e):
					print(f'{type(e)}: {str(e)}')
				break

		#
		# Values
		#
		r_key = winreg.OpenKey(r_conn, path)
		i = 0
		while True:
			try:
				name, value, type_ = winreg.EnumValue(r_key, i)

				result['values'][name] = value

				i += 1
			except OSError as e:
				if '[WinError 259]' not in str(e):
					print(f'{type(e)}: {str(e)}')
				break

		return result

	def fix_uninstall_cmd(self, cmd:str):
		"""
		C:\Program Files (x86)\Whatever\installer.exe /remove /quiet
		->
		"C:\Program Files (x86)\Whatever\installer.exe" /remove /quiet
		"""
		cmd = re.sub('^([^/]+)(\s+/[a-zA-Z/\s]+)?$', '"$1"$2', cmd)

		return cmd


	#######                  #####
	#       # #      ###### #     # #   #  ####  ##### ###### #    #       #####    ##    ####  ###### #####
	#       # #      #      #        # #  #        #   #      ##  ##       #    #  #  #  #      #      #    #
	#####   # #      #####   #####    #    ####    #   #####  # ## #       #####  #    #  ####  #####  #    #
	#       # #      #            #   #        #   #   #      #    #       #    # ######      # #      #    #
	#       # #      #      #     #   #   #    #   #   #      #    #       #    # #    # #    # #      #    #
	#       # ###### ######  #####    #    ####    #   ###### #    #       #####  #    #  ####  ###### #####
	def expand_path(self, path):
		for (key, value) in os.environ.items():
			path = re.sub(
				f'%{key}%',
				value.replace('\\', '\\\\'),
				path,
				flags=re.IGNORECASE
			)

		return path

	def exists(self, path):
		path = self.expand_path(path)
		return os.path.exists(path)

	def chdir(self, path:str):
		path = self.expand_path(path)
		os.chdir(path)

	def copy(self, source:str, destination_dir:str):
		source = self.expand_path(source)
		destination_dir = self.expand_path(destination_dir)

		source = source.rstrip('\\/')
		destination_dir = destination_dir.rstrip('\\/')

		if not os.path.exists(destination_dir):
			os.makedirs(destination_dir)

		source_mask = '^' + re.escape(source.rstrip('*'))\
			.rstrip('\\/')\
			.replace('\\*\\*', '.{0,}')\
			.replace('\\*', '[^\\\\/]*')

		for item in glob.glob(source, recursive=True):
			item_relative = re.sub(source_mask, '', item).strip('\\/')
			new_item = f'{destination_dir}\\{item_relative}'

			if os.path.exists(new_item):
				continue

			if os.path.isfile(item):
				shutil.copy(item, new_item)
			elif os.path.isdir(item):
				os.mkdir(new_item)

	def has_desktop_shortcut(self, name:str):
		name = re.sub(r'\.lnk$', '', name)

		folders = [
			r'%UserProfile%\Desktop',
			r'C:\Users\Public\Desktop',
		]
		for folder in folders:
			if self.exists(rf'{folder}\{name}.lnk'):
				return True

		return False

	def clear_desktop(self, keep=[]):
		self.clear_folder(r'%UserProfile%\Desktop', keep=keep)
		self.clear_folder(r'C:\Users\Public\Desktop', keep=keep)

	def clear_folder(self, folder, keep=[]):
		folder = self.expand_path(folder)

		files = os.listdir(folder)
		for file in files:
			# Check for `keep` masks
			delete = True
			for mask in keep:
				if re.match(rf'.*{mask}.*', file, re.IGNORECASE):
					delete = False
					break

			# Delete file
			if delete:
				path = f'{folder}\\{file}'
				try:
					os.unlink(path)
				except PermissionError:
					print(f'! PermissionError: unlink({path})')

	def find_uninstall_script_in_files(self, mask):
		paths = [
			fr'C:\Program Files\{mask}\uninst*.exe',
			fr'C:\Program Files (x86)\{mask}\uninst*.exe',
		]

		found = []
		for path in paths:
			for file in glob.glob(path, recursive=True):
				found.append(file)

		return found


	#     #
	##   ##   ##   ##### #####       #####    ##    ####  ###### #####
	# # # #  #  #    #     #         #    #  #  #  #      #      #    #
	#  #  # #    #   #     #         #####  #    #  ####  #####  #    #
	#     # ######   #     #         #    # ######      # #      #    #
	#     # #    #   #     #         #    # #    # #    # #      #    #
	#     # #    #   #     #         #####  #    #  ####  ###### #####
	def setup_taskbar_search(self, mode):
		if not self.is_win_10():
			return

		self.matt.right_click('taskbar')
		self.matt.click('taskbar_option_search')

		modes = {
			'hidden': 'taskbar_option_search_hidden',
			'icon': 'taskbar_option_search_icon',
			'box': 'taskbar_option_search_box',
		}
		self.matt.click(modes[mode])

	def setup_taskbar_news(self, mode):
		if not self.is_win_10():
			return

		self.matt.right_click('taskbar')
		self.matt.click('taskbar_option_news')

		modes = {
			'icon_text': 'taskbar_option_news_icon_text',
			'icon': 'taskbar_option_news_icon',
			'off': 'taskbar_option_news_off',
		}
		self.matt.click(modes[mode])

	def search(self, name:str):
		self.matt.hotkey('win', 's')
		self.matt.wait('search_ready')
		self.matt.typewrite(name)
		self.matt.wait('search_app_done')

	def unpin_from_taskbar(self, apps:list):
		for app in apps:
			try:
				self.search(app)
				self.matt.right_click('search_app_done')
				self.matt.click('unpin_from_taskbar', timeout=1.0)
			except TimeoutError:
				self.matt.hotkey('esc')

			self.matt.hotkey('ctrl', 'a')
			self.matt.hotkey('del')

		self.matt.hotkey('esc')

	def pin_to_taskbar(self, apps:list):
		for app in apps:
			try:
				self.search(app)
				self.matt.right_click('search_app_done')
				self.matt.click('pin_to_taskbar', timeout=1.0)
			except TimeoutError:
				self.matt.hotkey('esc')

			self.matt.hotkey('ctrl', 'a')
			self.matt.hotkey('del')

		self.matt.hotkey('esc')

	# def run(self, cmd:str):
	# 	# subprocess.run(['explorer.exe', 'ms-settings:regionlanguage'])
	# 	# subprocess.Popen(['explorer.exe', 'ms-settings:regionlanguage'])
	#
	# 	self.matt.hotkey('win', 'r')
	# 	self.matt.wait('run_title')
	# 	self.matt.typewrite(cmd)
	# 	self.matt.hotkey('enter')

	# def open_settings_language(self):
	# 	# W10: ms-settings:regionlanguage-setdisplaylanguage
	# 	# W11: ms-settings:regionlanguage                     # works on W10 as well
	# 	self.run('ms-settings:regionlanguage')

import os, sys, pathlib, re, platform, subprocess
from time import sleep

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from matt.matt import Matt


"""
https://winaero.com/ms-settings-commands-in-windows-10/
http://woshub.com/ms-settings-uri-commands-windows-11/
https://support.microsoft.com/en-us/windows/windows-keyboard-shortcuts-3d444b08-3a00-abd6-67da-ecfc07e86b98
"""
class Windows():
	CMD_DISABLE_TRANSPARENCY = 'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name EnableTransparency -Value 0 -Type Dword -Force'
	CMD_ACTIVATE_DARK_THEME_APPS = 'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name AppsUseLightTheme -Value 0 -Type Dword -Force'
	CMD_ACTIVATE_DARK_THEME_SYSTEM = 'Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize -Name SystemUsesLightTheme -Value 0 -Type Dword -Force'
	CMD_SETUP_LANGUAGES = 'Set-WinUserLanguageList -Force -LanguageList {}'
	CMD_SET_DISPLAY_LANGUAGE = 'Set-WinUILanguageOverride -Language {}'

	def __init__(self, cache_path:str) -> None:
		self.matt = Matt(cache_file=cache_path)

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

	def run_powershell_command(self, command):
		subprocess.call([
			r'%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe',
			command
		], shell=True)

	def is_win_10(self):
		# https://stackoverflow.com/questions/66554824/get-windows-version-in-python
		version = int( platform.version().split('.')[2] )

		return (19000 <= version <= 21999)

	def unify_ui(self, languages='en-US,cs-CZ', display_language='en-US'):
		self.run_powershell_command(
			f'{self.CMD_ACTIVATE_DARK_THEME_SYSTEM}; ' + \
			f'{self.CMD_ACTIVATE_DARK_THEME_APPS}; ' + \
			f'{self.CMD_DISABLE_TRANSPARENCY}; ' + \
			f'{self.CMD_SETUP_LANGUAGES.format(languages)}; ' + \
			f'{self.CMD_SET_DISPLAY_LANGUAGE.format(display_language)}; ' + \
			''
		)

	def disable_transparency(self):
		self.run_powershell_command(self.CMD_DISABLE_TRANSPARENCY)

	def activate_dark_theme(self):
		self.run_powershell_command(f'{self.CMD_ACTIVATE_DARK_THEME_APPS}; {self.CMD_ACTIVATE_DARK_THEME_SYSTEM}')

	def setup_languages(self, languages):
		self.run_powershell_command(self.CMD_SETUP_LANGUAGES.format(languages))

	def set_display_language(self, language):
		self.run_powershell_command(self.CMD_SET_DISPLAY_LANGUAGE.format(language))

	def clear_desktop(self, keep=[]):
		self.clear_folder(str(pathlib.Path.home()) + '\\Desktop', keep=keep)
		self.clear_folder('C:\\Users\\Public\\Desktop', keep=keep)

	def clear_folder(self, folder, keep=[]):
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

	def show_desktop_this_pc(self):
		pass
		# self.matt.hotkey('win', 'd')
		# self.run(r'%windir%\System32\rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0')

		# this_pc, pos = self.matt.which('desktop_icon_this_pc_active', 'desktop_icon_this_pc_inactive')
		# if this_pc == 'desktop_icon_this_pc_inactive':
		# 	self.matt.click('desktop_icon_this_pc_inactive')

		# self.matt.hotkey('enter')

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

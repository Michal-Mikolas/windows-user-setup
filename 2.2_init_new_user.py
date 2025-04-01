import os, glob, shutil
from windows.windows import Windows
import config
from datetime import datetime
from getpass import getpass

def log(msg):
	print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")

def cmd_result(cmd:list, join_lines=False):
	import subprocess

	result = ''
	try:
		result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
		error = result.stderr
		result = result.stdout

		if join_lines:
			lines = [l for l in result.strip().splitlines() if l.strip()]
			result = join_lines.join(lines)
	except:
		pass

	return result

def netlog():
	import os, requests

	# Get laptop info
	user = os.getlogin()
	laptop = os.environ.get("COMPUTERNAME", "Unknown")
	laptop += ' - ' + cmd_result(['wmic', 'csproduct', 'get', 'vendor,name,version', '/value'], ' | ')
	laptop += ' | ' + cmd_result(['wmic', 'bios', 'get', 'serialnumber', '/value'], ' | ')

	# Log
	url = ("https://script.google.com/macros/s/AKfycbzAAOFjqgH9CCa-RonJRD2hVOhlvv2Aad_iLtO0a0L3UQ-AFp7xcOojBVR8efPnjk69/exec"
		   f"?LAPTOP={laptop}&USER={user}")

	try:
		response = requests.get(url)
		response.raise_for_status()  # Raise an error for HTTP errors
		data = response.json()

		if data.get("status") == "success":
			return data  # Return full response if status is success
		else:
			print("API call failed:", data)
			exit() if getpass("Check your internet connection and try again... ") != 'Station123' else None
			return None

	except requests.exceptions.RequestException as e:
		print("Request failed:", e)
		exit() if getpass("Check your internet connection and try again... ") != 'Station123' else None
		return None

netlog()

laptop = os.environ.get("COMPUTERNAME", "Unknown")
if laptop[0:3] != 'NOT' and laptop[0:2] != 'PC':
	print(f'\n\nERROR: Laptop name is "{laptop}", rename the laptop and run this script again. \n\n')
	exit() if getpass(f'\n\nERROR: Laptop name is "{laptop}", rename the laptop and run this script again. ') != 'Station123' else None



#~~~~~~
#~~~~~~ Pombo installation
#~~~~~~

#
# 1. Copy Pombo to HDD
#
srcDir = '\\\\10.0.0.12\\all\\_INSTALL\\Software\\pombo'
destDir = 'C:\\Users\\Public'
pomboDir = 'C:\\Users\\Public\\pombo'

import os, shutil
from PyElevate import elevate
elevate()
import subprocess
subprocess.run(f'xcopy "{srcDir}" "{destDir}" /E/H/Y/R')

#
# 2. Make Pombo run on Windows startup
#
def run_schtasks(args):
	try:
		# Use shell=False and pass args as a list for better security and handling. Schtasks typically requires administrator privileges for system-wide tasks
		process = subprocess.run(['schtasks.exe'] + args,
								 capture_output=True,
								 text=True,
								 check=True,  # Raises CalledProcessError on non-zero exit code
								 creationflags=subprocess.CREATE_NO_WINDOW)  # Hide console window
		print(f"schtasks {args[0]} output:\n{process.stdout}")
		return True
	except FileNotFoundError:
		print("Error: schtasks.exe not found. Is it in your system's PATH?")
		return False
	except subprocess.CalledProcessError as e:
		print(f"Error running schtasks {' '.join(args)}.")
		print(f"Return Code: {e.returncode}")
		print(f"Output:\n{e.stdout}")
		print(f"Error Output:\n{e.stderr}")
		if "ERROR: Access is denied." in e.stderr or e.returncode == 5:
			print("\nHint: This command requires Administrator privileges.")
		return False
	except Exception as e:
		print(f"An unexpected error occurred running schtasks: {e}")
		return False

def create_startup_task(task_name, xml_path):
	print(f"\nAttempting to create scheduled task '{task_name}'...")

	if not os.path.exists(xml_path):
		print(f"Error: XML path not found: {xml_path}")
		return False
	if not os.path.isabs(xml_path):
		print(f"Error: Path must be absolute: {xml_path}")
		return False

	args = [
		'/create',
		'/tn', task_name,  # /tn: Task Name
		'/xml', xml_path,  # /xml: Path to the definition file
		'/f'               # /f: Force creation (overwrite if exists)
	]
	return run_schtasks(args)

def delete_startup_task(task_name):
	print(f"Attempting to delete scheduled task '{task_name}'...")
	args = ['/delete', '/tn', task_name, '/f']
	return run_schtasks(args)

def trigger_task(task_name):
	print(f"Attempting to run scheduled task '{task_name}' now...")

	# Construct the arguments for schtasks.exe /run
	args = [
		'/run',
		'/tn', task_name  # /tn: Task Name
	]

	# Call the helper function to execute the command
	success = run_schtasks(args)
	if success:
		print(f"Successfully requested Task Scheduler to run '{task_name}'.")
	else:
		print(f"Failed to run task '{task_name}'. Check errors above.")

	return success

task_created = create_startup_task('pombo', f"{pomboDir}\\pombo_task.xml")
startupDir = 'C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup'
if task_created:
	subprocess.run(f'rm "{startupDir}\\pombo.lnk"', shell=True)
else:
	subprocess.run(f'copy /Y "{pomboDir}\\pombo.lnk" "{startupDir}"', shell=True)

#
# 3. Run Pombo now
#
if task_created:
	trigger_task('pombo')
else:
	DETACHED_PROCESS = 0x00000008        # Prevents child process from closing when Python exits
	subprocess.Popen(
		f'{pomboDir}\\pombo.exe',
		cwd=pomboDir,
		creationflags=DETACHED_PROCESS,
		close_fds=True | subprocess.CREATE_NEW_PROCESS_GROUP,
		shell=True,
	)
#~~~~~~/
#~~~~~~/
#~~~~~~/



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
	log('- - Manage startup apps (keep enabled: ecmds.exe, Mattermost)')
	log('- - Restart PC')

except Exception as e:
	log(f'{type(e).__name__}: {str(e)}')

print('')
input('Press ENTER to continue...')

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


#~~~~~~
#~~~~~~ Pombo installation
#~~~~~~
def run_schtasks(args):
	import subprocess
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

def stop_task(task_name):
	print(f"Attempting to stop scheduled task '{task_name}' now...")

	# Construct the arguments for schtasks.exe /run
	args = [
		'/End',
		'/tn', task_name  # /tn: Task Name
	]

	# Call the helper function to execute the command
	success = run_schtasks(args)
	if success:
		print(f"Successfully requested Task Scheduler to stop '{task_name}'.")
	else:
		print(f"Failed to stop task '{task_name}'. Check errors above.")

	return success


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

	#
	log('- running apps for setup')
	apps = {
		'Edge': [
			'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
		],
		'Chrome': [
			'C:\\Program Files*\\Google\\Chrome\\Application\\chrome.exe',
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
					print('Install the Edge extension please!')
					print('')
					windows.run(f'"{file}" "https://microsoftedge.microsoft.com/addons/detail/datadeck/ahbodcpioanbmiaahimpflcopbckameo"')
					print('')

				if 'Chrome' in file:
					print('')
					print('Install the Chrome extension please!')
					print('')
					windows.run(f'"{file}" "https://chromewebstore.google.com/detail/datadeck/nebkjlepbimdhlfakkdkdnjgcgcplbmk"')
					print('')

				break

			if file:
				break

		if not file:
			print(f'- - - NOT FOUND')

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

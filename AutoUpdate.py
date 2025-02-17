import os
import re
import sys
import time
import json
import requests
import subprocess

os.system('color')

class COLOUR:
    ESC = '\x1b'
    GREEN  = ESC + '[32m'
    GREEN_BG  = ESC + '[42m'
    RED    = ESC + '[31m'
    RED_BG    = ESC + '[41m'
    YELLOW = ESC + '[33m'
    CYAN = "\033[0;36m"
    STOP   = '\x1b[0m'

def get_latest_version(url):
    try:
        response = requests.get(url)
        if response.status_code != 200:
            print(f"Status Code error: Status code {response.status_code}")
            return None
        files = response.json()
        
        # Sort files by name or any logic that determines the 'latest'
        version_files = []
        for file in files:
            match = re.search(r'v(\d+\.\d+(\.\d+)?)', file['name'])
            if match:
                version_files.append((file['name'], tuple(map(int, match.group(1).split('.')))))

        if version_files:
            latest_file = max(version_files, key=lambda x: x[1])[0]
            latest_version = re.search(r'v(\d+\.\d+(\.\d+)?)', latest_file).group(0)
        return latest_version  # Return the latest file name
    except requests.RequestException as e:
        print(f"Network error: {e}")
        return None
    except json.JSONDecodeError:
        print("Error decoding JSON response")
        return None

def get_local_version():
    # Extract the version from the file name of the executable
    exe_name = os.path.basename(__file__)
    match = re.search(r'v(\d+\.\d+(\.\d+)?)', exe_name)
    if match:
        local_version = match.group(0)  # e.g., v1.3.4 or v1.3 or v1
        return local_version
    else:
        print(f"{COLOUR.RED}Could not determine local version. Please check exe filename and report this bug to Kai.{COLOUR.STOP}")
        return None

def check_for_updates(local_version, latest_version, download_url):
    
    if latest_version and (latest_version != local_version):
        # Compare version numbers
        latest_version_num = tuple(map(int, latest_version.lstrip('v').split('.')))
        local_version_num = tuple(map(int, local_version.lstrip('v').split('.')))
        
        if latest_version_num > local_version_num:
            print(f"{COLOUR.YELLOW}New version available: {latest_version}{COLOUR.STOP}")
            # Download and install the update
            download_update(local_version, latest_version, download_url)
        else:
            print(f"{COLOUR.GREEN}You are using the latest version.{COLOUR.STOP}")
    else:
        print(f"{COLOUR.GREEN}You are using the latest version.{COLOUR.STOP}")

def download_update(local_version, latest_version, download_url):
    response = requests.get(download_url)
    if response.status_code == 200:
        new_file = f"batch_analyser_{latest_version}.exe"
        with open(new_file, "wb") as f:
            f.write(response.content)
        print("Update downloaded successfully.")
        
        delete = input(f"Close and delete current version {local_version}? (y/n): ").lower()

        current_exe = sys.argv[0]

        # Batch script that opens the new version
        bat_script = """
        @echo off
        timeout /t 2 >nul
        start "" "{new_file}"
        del "%~f0"
        """.format(new_file=new_file)

        if delete == "y":
            print(f"{COLOUR.RED}This will delete the current version from your system and is irreversible.{COLOUR.STOP}")
            confirm = input(f"Input 'delete' to confirm deletion of version {local_version}: ").lower()

            if confirm == "delete":
            # Batch script that deletes the current version, then opens the new version.
                bat_script = """
                @echo off
                timeout /t 2 >nul
                del "{current_exe}"
                start "" "{current_exe}"
                del "%~f0"
                """.format(current_exe=current_exe, new_file=new_file, current_exe_name=os.path.basename(current_exe))
            else:
                print(f"{COLOUR.YELLOW}Old version will not be deleted.{COLOUR.STOP}")
        
        bat_file = "update_script.bat"
        with open(bat_file, "w") as f:
            f.write(bat_script)
        
        print("Closing application for update...")
        time.sleep(2.0)
        subprocess.Popen(bat_file, shell=True)
        sys.exit(0)
    else:
        print("Failed to download the update.")
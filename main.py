import os
import win32api
import win32serviceutil
import subprocess
import time
import psutil
import datetime
import getpass
import platform
import json
from ScriptCounter import ScriptCounter
from openpyxl import load_workbook

# open the template workbook
wb = load_workbook('Test Script Workbook\\TEMPLATE Test Scripts Workbook - v0.10.xlsx')
# Select the worksheets by name
ws1 = wb["Pre-Test Checklist"]
ws2 = wb["Test Scripts"]
ws3 = wb["Post-Test Report"]
# Activate the writing to excel spreadsheet
ws = wb.active

# Loading the counting class for the test, for auditing reasons
count = ScriptCounter()
count = count.count

# This details the date and time from the machines clock
current_date = datetime.datetime.now().strftime(f"%Y-%m-%d")
current_time = datetime.datetime.now().strftime(f"%H:%M:%S")

# Creates directorys for the script to save documents and screenshots to
screenshotevidence = "Screenshot_Evidence"
workbook = "Test Script Workbook"
logfile = "Logs"

# Makes a Log folder in the script directory if one doesn't exist
if not os.path.exists(logfile):
    os.makedirs(logfile)

# Makes a Workbook folder in the script directory if one doesn't exist
if not os.path.exists(workbook):
    os.makedirs(workbook)

# Makes a Attempt folder in the workbook directory if one doesn't exist
if not os.path.exists(f"{workbook}/Attempt_{count}"):
    os.makedirs(f"{workbook}/Attempt_{count}")

# Makes a screenshot evidence foler in the Attempt directory if one doesn't exist
if not os.path.exists(f"{workbook}/Attempt_{count}/{screenshotevidence}"):
    os.makedirs(f"{workbook}/Attempt_{count}/{screenshotevidence}")

# Constructs the full path to the workbook file
workbook_filepath = os.path.join(os.getcwd(), workbook, f"Attempt_{count}", f"{platform.node()}_Test_Scripts_Workbook_{current_date}.xlsx")

# A function to get the file version of a given file
def get_file_version(file_path):
    # Attempt to retrieve the file version information from the given file
    try:
        # Use the win32api module to get the file version info
        info = win32api.GetFileVersionInfo(file_path, "\\")
        # Extract the "most significant" and "least significant" parts of the version number
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        # Combine the parts of the version number to form a string in the format "x.y.z.w"
        version = f'{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}'
        # Return the version string
        return version
    # If an error occurs (e.g. the file does not exist or does not contain version information), return "unknown"
    except Exception:
        return 'does not match,'

# A function to check if a service is running
def is_service_running(service_name):
    # Use the win32serviceutil module to query the status of the service
    status = win32serviceutil.QueryServiceStatus(service_name)
    # Return True if the status of the service is "running" (i.e. status code 4)
    return status[1] == 4

# A function to check if a process is running
def is_running(process_name):
    for proc in psutil.process_iter():
        try:
            if process_name.lower() in proc.name().lower():
                return True
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass
    return False

# The run_powershell_script function takes a file path as input, runs the specified PowerShell script using the subprocess module, and returns a boolean value based on the script's output.
def run_powershell_script(script_path):
    command = ['powershell', '-ExecutionPolicy', 'Unrestricted', '-File', script_path]
    output = subprocess.check_output(command)
    result = None
    try:
        result = json.loads(output.decode('utf-8').strip())
    except json.JSONDecodeError:
        pass
    return result

# Call the run_powershell_script function with the path to the PowerShell script
bitlocker_script_path = "check_bitlocker.ps1"
bitlocker_script_result = run_powershell_script(bitlocker_script_path)

# Access the results in the returned hashtable and handle any errors
if bitlocker_script_result is not None:
    bitlocker_encrypted = bitlocker_script_result.get('BitlockerEncrypted', False)
    bitlocker_service_running = bitlocker_script_result.get('BitlockerServiceRunning', False)
    mbam_service_running = bitlocker_script_result.get('MBAMServiceRunning', False)
else:
    bitlocker_encrypted = False
    bitlocker_service_running = False
    mbam_service_running = False

# Define a dictionary of application information, including the name, file=file path, reported version, and expected version this dictionary is used for AP items 1 - 65
applications = {
    "AP - 1": {
        "name": "Discord",
        "processname": "discord.exe",
        "path": 'C:\\Users\\jack_\\AppData\\Local\\Discord\\Update.exe',
        "reportedversion": get_file_version('C:\\Users\\jack_\\AppData\\Local\\Discord\\Update.exe'),
        "expectedversion" : "9.1.1.1"
    },
    "AP - 2": {
        "name": "Firefox",
        "processname": "firefox.exe",
        "path": 'C:\\Program Files\\Mozilla Firefox\\firefox.exe',
        "reportedversion": get_file_version('C:\\Program Files\\Mozilla Firefox\\firefox.exe'),
        "expectedversion" : "108.0.1.8384"
    },
    "AP - 3": {
        "name": "FireFox",
        "processname": "notepad.exe",
        "path": 'C:\\Program Files\\Mozilla Firefox\\fire.exe',
        "reportedversion": get_file_version('C:\\Program Files\\Mozilla Firefox\\fire.exe'),
        "expectedversion" : "1.1.1.0"
    },
    "AP - 4": {
        "name": "Anarchy Online",
        "processname": "Anarchy.exe",
        "path": 'C:\\Funcom\\Anarchy Online\\Anarchy.exe',
        "reportedversion": get_file_version('C:\\Funcom\\Anarchy Online\\Anarchy.exe'),
        "expectedversion" : "1.9.1.1"
    },
    "AP - 5": {
        "name": "notepad",
        "processname": "notepad.exe",
        "path": 'C:\\WINDOWS\\system32\\notepad.exe',
        "reportedversion": get_file_version('C:\\WINDOWS\\system32\\notepad.exe'),
        "expectedversion" : '10.0.19041.1865',
    }
}

# Define a dictionary of service names
services = {
    "OS - 10": "Parental Controls",
    "OS - 11": "DHCP Client",
    "OS - 12": "Network Virtualization Service",
    "OS - 13": "notrealservice",
}

Startupaplications = {
    "OS - 06": "Acrotray",
    "OS - 06": "Adobe Updater Startup Utility",
    "OS - 06": "AGCInvokerUtility",
    "OS - 06": "Box",
    "OS - 06": "Box Edit",
    "OS - 06": "Citrix Connection Center",
    "OS - 06": "Citrix FTA, URL Redirector",
    "OS - 06": "DisplayLink UI Syst-Tray Application",
    "OS - 06": "Microsoft OneDrive",
    "OS - 06": "Microsoft To Do",
    "OS - 06": "Program",
    "OS - 06": "Windows Security notification icon",
    "OS - 06": "Winzip Preloader",

}

# A dictionary for test OS - 04 
desktop_shortcuts = {
        "OS - 04": {
            "name": "",
            "path": 'C:\\users\\Public\\Desktop\\',
        },
        "OS - 04.1": {
            "name": "",
            "path": 'C:\\users\\Public\\Desktop\\,
        },
        "OS - 04.2": {
            "name": "",
            "path": 'C:\\users\\Public\\Desktop\\',
        }
}

# Opens a new notepadd file in the specific directory and writes to it
with open("Logs\\Log.txt", "w") as file:

    # Prints out how many times this script has been run
    print("This script has been run {} times".format(count), file=file)

    # This is the variables to determine how many tests passed or have failed and how many tests were taken.
    # Each variable is added to during an if statement.
    passed = 0
    failed = 0
    notrun = 0
    tests = 0


    # beginning the test and adding relevant details.
    print("-----------------------------------------", file=file)
    print("Beginning Test Scrip Workbook", file=file)
    print("-----------------------------------------", file=file)
    

    # This details the date and time from the machines clock
    print("Date:", current_date, file=file)
    print("Time:", current_time, file=file)

    # prints the windows version
    print("Windows version: ", platform.uname().version, file=file)

    # This gets the FQDN from the machine.
    print("Domain Name:", platform.node(), file=file)
        
    # Prints the user account that's running the script.
    print("Account Name:", getpass.getuser(), file=file)
    print("-----------------------------------------", file=file)
    print("Script created by Jack Gallacher.", file=file)
    print(" ", file=file)
    print(" ", file=file)
    print("-----------------------------------------", file=file)
    print("Beginning the C:\ drive encryption, Bitlocker service and MBAM Agent check.", file=file)
    try:
        # Statement to check if the C: Drive is encrypted 
        if bitlocker_encrypted:
            # Add lines to log file
            print("C: drive is encrypted", file=file)
            print("SE-01 is a PASS.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            passed += 1
            tests += 1
        else:
            # Add lines to log file
            print("C: drive is not encrypted", file=file)
            print("SE-01 is a Fail.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            failed += 1
            tests += 1
    except Exception as e:
        # Add lines to log file
        print("An error occurred:", e, file=file)
        print("Adding to the notrun value.", file=file)
        # Add inputs to excel file

        # Contribute to the stat variables
        tests += 1
        notrun = 1
    try:
    # Statement to check if the bitlocker service is running
        if bitlocker_service_running:
            # Add lines to log file
            print("BitLocker service is running", file=file)
            print("AP - 28 is a PASS.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            passed += 1
            tests += 1
        else:
            # Add lines to log file
            print("BitLocker service is not running", file=file)
            print("AP - 28 is a Fail.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            failed += 1
            tests += 1
    except Exception as e:
        # Add lines to log file
        print("An error occurred:", e, file=file)
        print("Adding to the notrun value.", file=file)
        # Add inputs to excel file

        # Contribute to the stat variables
        tests += 1
        notrun = 1
    try:
    # Statement to check if the bitlocker service is running
        if mbam_service_running:
            # Add lines to log file
            print("MBAM service is running", file=file)
            print("AP - 29 is a PASS.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            passed += 1
            tests += 1
        else:
            # Add lines to log file
            print("MBAM service is not running", file=file)
            print("AP - 29 is a Fail.", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            failed += 1
            tests += 1
    except Exception as e:
        # Add lines to log file
        print("An error occurred:", e, file=file)
        print("Adding to the notrun value.", file=file)
        # Add inputs to excel file

        # Contribute to the stat variables
        tests += 1
        notrun = 1

    # Add lines to log file    
    print("Ending the C:\ drive encryption, Bitlocker service and MBAM Agent check.", file=file)
    print("-----------------------------------------", file=file)
    print(" ", file=file)
    print(" ", file=file)
    print("-----------------------------------------", file=file)
    print("Beginning AP checks", file=file)


    # Iterate over the application information dictionary
    for workbook_entry, app_info in applications.items():

        # Get the version of the file at the given path
        version = get_file_version(app_info["path"])

        # Check if the path exists and if the version of the file at that path matches the expected version
        bothchecks = os.path.exists(app_info["path"]) and version == app_info["expectedversion"]

        # Check if the path exists and if the version of the file does not equal the expected version
        pathnoversion = os.path.exists(app_info["path"]) and version != app_info["expectedversion"]

        # Check if the path exists and if the version of the file doesn't exist 
        pathonly = os.path.exists(app_info["path"]) and app_info["expectedversion"] == None

        try:

            if bothchecks:
                # If both conditions are true, print a message indicating that the test is a pass
                # Add lines to log file
                print(f"Application Name: {app_info['name']} | version: {app_info['reportedversion']} exists.", file=file)
                print(f"This path exists: {app_info['path']}.", file=file)

                try:

                    # Opens the application and prints out that it has done so
                    if subprocess.Popen({app_info['path']}) != 0:   
                        # Add lines to log file
                        print(f"Opening {app_info['processname']}. ", file=file)
                        # Add inputs to excel file

                        # Contribute to the stat variables
                        passed += 1
                        tests += 1
                               
                    else:
                        # Add lines to log file
                        print(f"Can't open {app_info['path']} ", file=file)
                        # Add inputs to excel file

                        # Contribute to the stat variables
                        tests += 1
                        failed += 1


                # If it can't open the application, throw the error and say it can't open it           
                except Exception as e:
                    # Add lines to log file
                    print("An error occurred:", e, file=file)
                    print(f"Can't open {app_info['path']} ", file=file)
                    print(" ", file=file)
                    # Add inputs to excel file

                    # Contribute to the stat variables
                    tests += 1
                    notrun = 1

                # Wait for The application to open 
                if is_running(app_info['processname']):
                    # Add lines to log file
                    print(f"{app_info['processname']} has been opened succesfully.", file=file)
                    [p.kill() for p in psutil.process_iter() if p.name() == app_info['processname']]
                    print(f"{app_info['processname']} has been closed succesfully.", file=file) 
                    print(f"({workbook_entry}) is a PASS.", file=file)
                    print(" ", file=file)
                    # Add inputs to excel file

                    # Contribute to the stat variables
                    passed += 1
                    tests += 1
                    
                else:
                    # Add lines to log file
                    print(f"{app_info['name']} has not been opened succesfully.", file=file) 
                    print(f"({workbook_entry}) is a FAIL.", file=file)
                    print(" ", file=file)
                    # Add inputs to excel file

                    # Contribute to the stat variables
                    failed += 1
                    tests += 1
                    

            elif pathonly:
                # If both conditions are true, print a message indicating that the test is a pass, this will be used for shortcuts/links that have no version
                # Add lines to log file
                print(f"Application Name:{app_info['name']} | version doesn't exist.", file=file)
                print(f"This path exists: {app_info['path']}.", file=file)
                print(f"({workbook_entry}) is a PASS.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                passed += 1
                tests += 1        
                print(" ", file=file)

            elif pathnoversion:
                # If both conditions are true, print a message indicating that the test is a pass
                # Add lines to log file
                print(f"Expected version {app_info['expectedversion']} | actual version {app_info['reportedversion']}.", file=file)
                print(f"This path exists: {app_info['path']}.", file=file)
                print(f"({workbook_entry}) is a FAIL.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                failed += 1
                tests += 1


            else:
                # If either condition is not true, print a message indicating that the test is a fail
                # Add lines to log file
                print(f"Application Name: {app_info['name']} | version {app_info['reportedversion']}.", file=file)
                print(f"This path doesn't exists: {app_info['path']}.", file=file)
                print(f"({workbook_entry}) is a FAIL.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                failed += 1
                tests += 1
                

        except Exception as e:
            # Add lines to log file
            print("An error occurred:", e, file=file)
            print("Adding to the notrun value.", file=file)
            print(" ", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            tests += 1
            notrun = 1

    # Add lines to log file
    print("Ending AP checks", file=file)
    print("-----------------------------------------", file=file)
    print(" ", file=file)

    # Iterate over the service names dictionary
    for workbook_entry, services_info in services.items():  
        try:
            # Use the win32serviceutil module to query the status of the service
            # If the service is running (i.e. status code 4), print a message indicating that the service is running
            if win32serviceutil.QueryServiceStatus(services_info)[0] == 4 or is_service_running(services_info):

                # Add lines to log file
                print(f"({workbook_entry}), {services_info} exists and is running.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                passed += 1 
                tests += 1

            # If the service is not running, print a message indicating that the service is not running
            else:
                # Add lines to log file
                print(f"({workbook_entry}), {services_info} is not running.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                failed += 1
                tests += 1

        # If an error occurs (e.g. the service does not exist), print a message indicating that the service does not exist
        except Exception as e:
            # Add lines to log file
            print("An error occurred:", e, file=file)
            print(f"({workbook_entry}), {services_info} does not exist.", file=file)
            print(" ", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            failed += 1
            tests += 1
        

    # Iterate over the desktop shortcut dictionary 
    for workbook_entry, shortcut_info in desktop_shortcuts.items():
        try:
            if os.path.exists(shortcut_info["path"]):
                # Add lines to log file
                print("OS - 04 is a PASS", file=file)
                print(f"desktop shortcut name: {shortcut_info['name']}", file=file)
                print(f"This path exists: {shortcut_info['path']}.", file=file)
                print(" ")
                # Add inputs to excel file

                # Contribute to the stat variables
                passed += 1 
                tests += 1
            else:
                # Add lines to log file
                print("OS - 04 is a FAIL", file=file)
                print(f"desktop shortcut Name: {shortcut_info['name']}", file=file)
                print(f"This path doesn't exists: {shortcut_info['path']}.", file=file)
                print(" ", file=file)
                # Add inputs to excel file

                # Contribute to the stat variables
                failed += 1
                tests += 1
        except Exception as e:
            # Add lines to log file
            print("An error occurred:", e, file=file)
            print(f"Can't open {app_info['path']} ", file=file)
            print(" ", file=file)
            # Add inputs to excel file

            # Contribute to the stat variables
            tests += 1
            notrun = 1
        print(" ", file=file)


    # printing stats at the bottom fo the log file
    print("-----------------------------------------", file=file)
    print("Post Test Report", file=file)
    print("Windows version: ", platform.uname().version, file=file)
    print("Domain Name:", platform.node(), file=file)
    print(f"Number of tests passed:{passed}", file=file)
    print(f"Number of tests failed:{failed}", file=file)
    print(f"Number of tests ran: {tests}", file=file)
    print(f"Number of test no run: {notrun}", file=file)
    print("-----------------------------------------", file=file)


    # inputing figures from specific variables into the Post-Test Report section
    ws3['B3'] = f"{platform.uname().version}"
    ws3['B2'] =  f"{platform.node()}"
    ws3['B4'] = f"{passed}"
    ws3['B5'] = f"{failed}"
    ws3['B7'] = f"{tests}"
    ws3['B6'] = f"{notrun}"
    

        # Saves the Excel file to its attempt directory
    try:
        print(f"Saving workbook as {workbook_filepath}", file=file)
        wb.save(workbook_filepath)
    except Exception as e:
        print("An error occurred:", e, file=file)
        print("Can't save work book. Maybe close template or run as administrator", file=file)

# Close the file log file
file.close()


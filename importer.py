# -*- coding: utf-8 -*-
"""Nonprofit Salesforce © 2022 by 501 Commons is licensed under CC BY 4.0"""

"""import Module for Excel to Salesforce"""

"""Helpful Link: https://pbpython.com/windows-com.html"""

try:
    from win32api import STD_INPUT_HANDLE
    from win32console import GetStdHandle, ENABLE_PROCESSED_INPUT
except ImportError as ex:
    print str(ex)

class KeyboardHook():
    """Keyboard Hook Class"""

    def __enter__(self):
        self.readHandle = GetStdHandle(STD_INPUT_HANDLE)
        self.readHandle.SetConsoleMode(ENABLE_PROCESSED_INPUT)

        self.input_lenth = len(self.readHandle.PeekConsoleInput(10000))

        return self

    def __exit__(self, type, value, traceback):
        pass

    def reset(self):
        self.input_lenth = len(self.readHandle.PeekConsoleInput(10000))

        return True

    def key_pressed(self):
        """poll method to check for keyboard input"""

        events_peek = self.readHandle.PeekConsoleInput(10000)

        #Events come in pairs of KEY_DOWN, KEY_UP so wait for at least 2 events
        if len(events_peek) >= (self.input_lenth + 2):
            self.input_lenth = len(events_peek)
            return True

        return False

def main():
    """Main entry point"""

    import sys
    import os
    from os import listdir, makedirs
    from os.path import exists, join

    #
    # Required Parameters
    #

    salesforce_type = str(sys.argv[1])
    client_type = str(sys.argv[2])
    client_subtype = str(sys.argv[3])
    client_emaillist = str(sys.argv[4])

    if len(sys.argv) < 5:
        print ("Calling error - missing required inputs.  Expecting " +
               "salesforce_type client_type client_subtype client_emaillist\n")
        return

    print ("\nIncoming required parameters: " +
           "salesforce_type: {} client_type: {} client_subtype: {} client_emaillist: {} sys.argv {}\n"
           .format(salesforce_type, client_type, client_subtype, client_emaillist, sys.argv))

    print ("\n\nWhen import complete a status email with be sent to {}\n\n"
           .format(client_emaillist))

    print ("\n\nThis process can take up to 30 minutes to complete...")

    #
    # Optional Parameters
    #

    wait_time = 300
    if '-waittime' in sys.argv:
        wait_time = int(sys.argv[sys.argv.index('-waittime') + 1])

    norefresh = False
    if '-norefresh' in sys.argv:
        norefresh = True

    noupdate = False
    if '-noupdate' in sys.argv:
        noupdate = True

    enabledelete = False
    if '-enabledelete' in sys.argv:
        enabledelete = True

    noexportodbc = False
    if '-noexportodbc' in sys.argv:
        noexportodbc = True

    noexportsf = False
    if '-noexportsf' in sys.argv:
        noexportsf = True

    global emailattachments
    emailattachments = False
    if '-emailattachments' in sys.argv:
        emailattachments = True

    interactivemode = False
    if '-interactivemode' in sys.argv:
        interactivemode = True

    displayalerts = False
    if '-displayalerts' in sys.argv:
        displayalerts = True

    skipexcelrefresh = False
    if '-skipexcelrefresh' in sys.argv:
        skipexcelrefresh = True

    insert_attempts = 10
    if '-insertattempts' in sys.argv:
        insert_attempts = int(sys.argv[sys.argv.index('-insertattempts') + 1])

    location_local = True
    if 'Cloud' in sys.argv:
        location_local = False

    importer_root = ("C:\\repo\\Salesforce-Importer-Private\\Clients\\" + client_type +
                     "\\Salesforce-Importer")
    if '-rootdir' in sys.argv:
        importer_root = sys.argv[sys.argv.index('-rootdir') + 1]

    # Setup Logging to File
    sys_stdout_previous_state = sys.stdout
    if not interactivemode:
        sys.stdout = open(join(importer_root, '..\\importer.log'), 'w')
    print 'Importer Startup'

    importer_directory = join(importer_root, "Clients\\" + client_type)
    print "Setting Importer Directory: " + importer_directory

    # Global to monitor if should exit all processing
    global stop_processing
    stop_processing = False

    #Cloud location setup status results
    if not location_local:

        f = open(join(importer_directory, "ImportInstance_Status.txt"), "w")
        f.write("Complete")
        f.close()

    #Clear out log directory
    importer_log_directory = join(importer_root, "..\\Status\\")
    print "Check Status Directory: " + importer_log_directory
    if not exists(importer_log_directory):
        makedirs(importer_log_directory)

    importer_log_directory = join(importer_log_directory, client_subtype)
    print "Check Status Client Directory: " + importer_log_directory
    if not exists(importer_log_directory):
        makedirs(importer_log_directory)

    print "Clearing out the Importer Log Directory: " + importer_log_directory
    for file_name_only in listdir(importer_log_directory):
        file_name_full = join(importer_log_directory, file_name_only)
        if os.path.isfile(file_name_full):
            os.remove(file_name_full)

    # Export External Data
    status_export = ""

    if not noexportodbc:
        print "\n\nExporter - Export External Data\n\n"
        status_export = export_odbc(importer_directory,
                                    salesforce_type,
                                    client_subtype,
                                    interactivemode,
                                    displayalerts)

    # Check filename for operation
    insertOnly = False
    if "insert" in client_subtype.lower():
        insertOnly = True

    updateOnly = False
    if "update" in client_subtype.lower() or "upsert" in client_subtype.lower():
        updateOnly = True

    reportOnly = False
    if "report" in client_subtype.lower() and not updateOnly and not insertOnly:
        reportOnly = True

    print "norefresh: " + str(norefresh)
    print "noupdate: " + str(noupdate)
    print "insertOnly: " + str(insertOnly)
    print "updateOnly: " + str(updateOnly)
    print "reportOnly: " + str(reportOnly)

    # Insert Data
    status_import = ""
    if not norefresh and not updateOnly and not reportOnly and "Invalid Return Code" not in status_export:
        for insert_run in range(0, insert_attempts):

            print "\n\nImporter - Insert Data Process (run: %d)\n\n" % (insert_run)

            status_import = process_data(importer_directory, salesforce_type, client_type,
                                         client_subtype, 'Insert', wait_time,
                                         noexportsf,
                                         interactivemode,
                                         displayalerts,
                                         skipexcelrefresh,
                                         location_local)

            if stop_processing:
                return

            # Insert files are empty so continue to update process
            if "import_dataloader (returncode)" not in status_import:
                break

    # Update Data
    if not noupdate and not insertOnly and not reportOnly and not contains_error(status_import):
        print "\n\nImporter - Update Data Process\n\n"

        status_import = process_data(importer_directory, salesforce_type, client_type,
                                    client_subtype, 'Upsert', wait_time,
                                    noexportsf,
                                    interactivemode,
                                    displayalerts,
                                    skipexcelrefresh,
                                    location_local)

        status_import += process_data(importer_directory, salesforce_type, client_type,
                                     client_subtype, 'Update', wait_time,
                                     noexportsf, interactivemode, displayalerts, skipexcelrefresh, location_local)

    # Report Data
    if reportOnly and not insertOnly and not updateOnly:
        print "\n\nImporter - Report Data Process\n\n"

        status_import += process_data(importer_directory, salesforce_type, client_type,
                                     client_subtype, 'Report', wait_time,
                                     noexportsf, interactivemode, displayalerts, skipexcelrefresh, location_local)

    if stop_processing:
        return

    # Delete Data
    if enabledelete and not insertOnly and not updateOnly and not contains_error(status_import):
        print "\n\nImporter - Delete Data Process\n\n"
        status_import = process_data(importer_directory, salesforce_type, client_type,
                                     client_subtype, 'Delete', wait_time,
                                     noexportsf, interactivemode, displayalerts, skipexcelrefresh, location_local)

    if stop_processing:
        return

    # Restore stdout
    sys.stdout = sys_stdout_previous_state

    output_log = ""
    if not interactivemode:
        with open(join(importer_root, "..\\importer.log"), 'r') as exportlog:
            output_log = exportlog.read()

    file_path = importer_directory + "\\Status"
    import datetime
    date_tag = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    with open(join(file_path, "Salesforce-Importer-Log-{}.txt".format(date_tag)),
              "w") as text_file:
        text_file.write(output_log)

    #Write log to stdout
    print output_log

    if contains_error(status_import):
        #Cloud location setup status results
        if not location_local:
            f = open(join(importer_directory, "ImportInstance_Status.txt"), "w")
            f.write("Complete_With_Errors")
            f.close()

    # Send email results
    results = "Success"
    if contains_error(status_import) or contains_error(status_export):
        results = "Error"
    subject = "{}-{} Salesforce Importer Results - {}".format(client_type, client_subtype, results)

    try:
        send_email(client_emaillist, subject, file_path, emailattachments, importer_log_directory)
    except Exception as ex:
        print "\nsend_email - Unexpected send email error:" + str(ex)

    print "\nImporter process completed\n"

def process_data(importer_directory, salesforce_type, client_type,
                 client_subtype, operation, wait_time,
                 noexportsf, interactivemode, displayalerts, skipexcelrefresh, location_local):
    """Process Data based on operation"""

    #Create log file for import status and reports
    from os import makedirs
    from os.path import exists, join
    file_path = importer_directory + "\\Status"
    if not exists(file_path):
        makedirs(file_path)

    output_log = "Process Data (" + operation + ")\n\n"
    status_process_data = ""

    # Export data from Salesforce

    try:
        if not noexportsf:
            status_process_data = export_dataloader(importer_directory,
                                                    salesforce_type, interactivemode, displayalerts, location_local, client_type, client_subtype)
        else:
            status_process_data = "Skipping export from Salesforce"
    except Exception as ex:
        output_log += "\n\nexport_dataloader - Unexpected error:" + str(ex)
        output_log += "\n\export_dataloader\n" + status_process_data
        status_process_data = "Error detected so skip processing - Exception"
    else:
        output_log += "\n\nExport\n" + status_process_data

    global stop_processing
    if stop_processing:
        return ""

    # Export data from Excel

    try:
        if (not skipexcelrefresh and not contains_error(status_process_data)
                and not contains_error(output_log.lower())):
            status_process_data = refresh_and_export(importer_directory,
                                                     salesforce_type, client_type,
                                                     client_subtype, operation,
                                                     wait_time, interactivemode, displayalerts)
        else:
            status_process_data = "Skipping refresh and export from Excel"
    except Exception as ex:
        output_log += "\n\nrefresh_and_export - Unexpected error:" + str(ex)
        output_log += "\n\refresh_and_export\n" + status_process_data
        status_process_data = "Error detected so skip processing - Exception"
    else:
        output_log += "\n\nExport\n" + status_process_data

    # Import Data into Salesforce

    if not "report" == operation.lower():

        try:
            if not contains_error(status_process_data) and not contains_error(output_log):
                status_process_data = import_dataloader(importer_directory,
                                                        client_type, salesforce_type,
                                                        operation)
            else:
                print status_process_data + output_log
                status_process_data = "Error detected so skip processing"
        except Exception as ex:
            output_log += "\n\nrefresh_and_export - Unexpected error:" + str(ex)
            output_log += "\n\import_dataloader\n" + status_process_data
            status_process_data = "Error detected so skip processing - Exception"
        else:
            output_log += "\n\nImport\n" + status_process_data

    import datetime
    date_tag = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    with open(join(file_path, "Salesforce-Importer-Log-{}-{}.txt".format(operation, date_tag)),
              "w") as text_file:
        text_file.write(output_log)

    return status_process_data + output_log

def open_workbook(xlapp, xlfile):
    try:        
        xlwb = xlapp.Workbooks(xlfile)            
    except Exception as e:
        try:
            # workbooks.open(file, UpdateLinks = No, ReadOnly = True, Format = 2 Commas)
            xlwb = xlapp.Workbooks.Open(xlfile, 0, True, 2)
        except Exception as e:
            print(e)
            xlwb = None                    
    return(xlwb)

def refresh_and_export(importer_directory, salesforce_type,
                       client_type, client_subtype, operation,
                       wait_time, interactivemode, displayalerts):
    """Refresh Excel connections"""

    import os
    import os.path
    import time
    import win32com.client as win32

    refresh_status = "refresh_and_export\n"

    excel_connection = win32.gencache.EnsureDispatch("Excel.Application")
    excel_connection.Visible = interactivemode

    excel_file_path = importer_directory + "\\"
    excel_file = excel_file_path + client_type + "-" + client_subtype + "_" + salesforce_type + ".xlsx"

    global workbook
    workbook_assigned = False
    workbook_successful = False
    open_max_attempts = 5
    open_attempt = 0
    found_operation_sheet = True

    while open_attempt < open_max_attempts and found_operation_sheet:

        open_wait_time = wait_time
        open_attempt += 1

        message = "\nImport Process - Attempt " + str(open_attempt) + " of " + str(open_max_attempts) + " to open Excel: " + excel_file
        print message
        if not os.path.exists(excel_file):
            message = "Import Process - ERROR File does not exist: " + excel_file
            print message

        try:
            workbook = open_workbook(excel_connection, excel_file)
            workbook_assigned = True

            found_operation_sheet = False
            for sheet in workbook.Sheets:
                sheet_name_lower = sheet.Name.lower()
                if operation.lower() in sheet_name_lower:
                    found_operation_sheet = True
                    break

            if not found_operation_sheet:
                refresh_status += "No sheets matched the operation: " + operation + "\n"
                print refresh_status

            else:

                message = "\nImport Process - Pausing 30 seconds for Excel to load in the background (You can see Excel in Task Manager but will be hidden from the desktop for better performance)..."
                print message
                refresh_status += message + "\n"
                time.sleep(30)

                #for connection in workbook.Connections:
                    #print connection.name
                    # BackgroundQuery does not work so have to do manually in Excel for each Connection
                    #connection.BackgroundQuery = False

                # RefreshAll is Synchronous iif
                #   1) Enable background refresh disabled/unchecked in xlsx for all Connections
                #   2) Include in Refresh All enabled/checked in xlsx for all Connections
                #   To verify: Open xlsx Data > Connections > Properties for each to verify
                message = "\nImport Process - Refreshing all connections..."
                print message
                refresh_status += message + "\n"

                # RefreshAll - if direct Salesforce connection then will prompt for username & password
                #       under a couple of scenarios and will block until creds updates
                #   Scenario 1: First time running automation on a particular machine.
                #       User needs to select Remember me or this Scenario will repeat
                #   Scenario 2: Salesforce Password changed
                #   Scenario 3: Excel I think has a 3 month expiration for the user cred cookie
                #
                # Avoid adding connections to Excel that require username/password
                #   (e.g., Salesforce, Database).
                #   Instead use Exporter to pull the data external to Excel.
                workbook.RefreshAll()

                # Wait for excel to finish refresh
                message = ("Pausing " + str(open_wait_time) +
                        " seconds to give Excel time to complete background query...")
        #                   "\n\t\t***if Excel background query complete then press any key to exit wait cycle")
                print message
                refresh_status += message + "\n"

        #        with KeyboardHook() as keyboard_hook:

                    #Clear the input buffer
        #            keyboard_hook.reset()

                while open_wait_time > 0:
                    if open_wait_time > 30:
                        time.sleep(30)

                        open_wait_time -= 30
                        message = ("\t" + str(open_wait_time) +
                                    " seconds remaining for Excel to complete background query...")
        #                               "\n\t\t***if Excel background query complete then press any key to exit wait cycle")
                        print message
                        refresh_status += message + "\n"

                    else:
                        time.sleep(open_wait_time)
                        open_wait_time = 0
                        break

        #                if keyboard_hook.key_pressed():
        #                    print "\nUser interrupted wait cycle\n"
        #                    break

                message = "Import Process - Refreshing all connections...Completed"
                print message
                refresh_status += message + "\n"

                if not os.path.exists(excel_file_path + "Import\\"):
                    os.makedirs(excel_file_path + "Import\\")

                update_sheet_found = False
                for sheet in workbook.Sheets:
                    sheet_name_lower = sheet.Name.lower()
                    if "update" in sheet_name_lower:
                        update_sheet_found = True
                        break

                for sheet in workbook.Sheets:

                    # Only export update, insert, upsert, delete, or report sheets
                    sheet_name_lower = sheet.Name.lower()
                    if ("update" not in sheet_name_lower
                            and "upsert" not in sheet_name_lower
                            and "insert" not in sheet_name_lower
                            and "delete" not in sheet_name_lower
                            and "report" not in sheet_name_lower):
                        continue

                    excel_connection.Sheets(sheet.Name).Select()
                    sheet_file = excel_file_path + "Import\\" + sheet.Name + ".csv"

                    message = "Exporting csv for sheet: " + sheet_file
                    print message
                    refresh_status += message + "\n"

                    # Save report to Status to get attached to email
                    if "report" in sheet.Name.lower():

                        # Check if Manifest meaning report needs to be split  up
                        if "manifest" in sheet.Name.lower():

                            sheet_file = ""
                            process_manifest(workbook, sheet.Name, excel_file_path + "Status\\")
                        else:
                            sheet_file = excel_file_path + "Status\\" + sheet.Name + ".csv"

                    # Check for existing file
                    if os.path.isfile(sheet_file):
                        os.remove(sheet_file)

                    # By Design - set displayalerts before saveas so not prompting w/ save dialogs during automation.  Moved this here so that any RefreshAll errors will still surface and cause the refresh process not to finish thus an error will be detected
                    excel_connection.DisplayAlerts = displayalerts

                    if not sheet_file == "":
                        workbook.SaveAs(sheet_file, 6)

                    # Update check to make sure insert sheet is empty
                    if (operation == "Update"
                            and update_sheet_found
                            and "insert" in sheet.Name.lower()
                            and contains_data(sheet_file)):

                        raise Exception("refresh_and_export: Update Error", (
                            "Insert sheet contains data and should be empty during update process: " +
                            sheet_file))

            workbook_successful = True

        except Exception as ex:
            message += "Unexpected error:" + str(ex)
            print message
            refresh_status += message + "\n"

            if open_attempt >= open_max_attempts:
                excel_connection.Quit()
                raise Exception("refresh_and_export", refresh_status)

            message = "\nImport Process - Pausing 30 seconds for system to recover from error..."
            print message
            refresh_status += message + "\n"
            time.sleep(30)

        finally:
            if not workbook is None and workbook_assigned:
                workbook.Close(False)

            workbook_assigned = False

            if workbook_successful:
                break;

    excel_connection.Quit()

    return refresh_status

# workbook details: https://learn.microsoft.com/en-us/office/vba/api/excel.workbook
def process_manifest(workbook, sheet_name, statusDirectory):

    import csv
    import pandas as pd
    print("The Version of Pandas is: ", pd.__version__)
    import sys
    import os
    import os.path
    from datetime import datetime

    sheetFile = os.path.join(statusDirectory, sheet_name, ".csv")

    print "process_manifest: " + sheetFile

    # Check for existing file
    if os.path.isfile(sheetFile):
        os.remove(sheetFile)

    workbook.SaveAs(sheetFile, 6)
    data = pd.read_csv(sheetFile)

    dateToday = datetime.today()

    for (cruiseID, cruiseDate), group in data.groupby(['Cruise ID', 'Cruise Date']):

        cruiseDateValue = datetime.strptime(cruiseDate, "%m/%d/%Y")
        daysDifference = abs((cruiseDateValue - dateToday).days)

        manifestType = "Preliminary"
        if daysDifference <= 10:
            manifestType = "Final"

        groupFileName = os.path.join(statusDirectory, "{}-{}-{}.csv".format(sheet_name, cruiseID, manifestType))
        group.to_csv(groupFileName, index=False)

    # Remove full sheet data file
    os.remove(sheetFile)

def contains_data(file_name):
    """Check if file contains data after header"""

    line_index = 1
    with open(file_name) as file_open:
        for line in file_open:
            # Check if line empty
            line_check = line.replace(",", "")
            line_check = line_check.replace('"', '')
            if (line_index == 2 and line_check != "\n"):
                return True
            elif line_index > 2:
                return True

            line_index += 1

    return False

def file_linecount(file_name):
    """Count how many lines after the header"""

    # set index to -1 so the header is not counted
    line_index = -1
    with open(file_name) as file_open:
        for line in file_open:
            if line:
                line_index += 1

    return line_index

def import_dataloader(importer_directory, client_type, salesforce_type, operation):
    """Import into Salesforce using DataLoader"""

    import os
    from os import listdir
    from os.path import join
    from subprocess import Popen, PIPE

    bat_path = importer_directory + "\\DataLoader"
    import_path = importer_directory + "\\Import"

    return_code = ""
    return_stdout = ""
    return_stderr = ""

    import_successful = False

    for file_name in listdir(bat_path):
        if not operation in file_name or ".sdl" not in file_name:
            continue

        # Check if associated csv has any data
        sheet_name = os.path.splitext(file_name)[0]
        import_file = join(import_path, sheet_name + ".csv")
        if not os.path.exists(import_file) or not contains_data(import_file):
            continue

        bat_file = (join(bat_path, "RunDataLoader.bat")
                    + " " + salesforce_type + " "  + client_type + " " + sheet_name)

        message = "Starting Import Process: " + bat_file + " for file: " + import_file
        print message
        return_stdout += message + "\n"
        import_process = Popen(bat_file, stdout=PIPE, stderr=PIPE)

        stdout, stderr = import_process.communicate()

        return_code += "import_dataloader (returncode): " + str(import_process.returncode)
        return_stdout += "\n\nimport_dataloader (stdout):\n" + stdout
        return_stderr += "\n\nimport_dataloader (stderr):\n" + stderr

        if (import_process.returncode != 0
                or contains_error(return_stdout)
                or "We couldn't find the Java Runtime Environment (JRE)" in return_stdout):
            raise Exception("Invalid Return Code", return_code + return_stdout + return_stderr)

        status_path = importer_directory + "\\status"

        for file_name_status in listdir(status_path):
            file_name_status_full = join(status_path, file_name_status)
            if contains_error(file_name_status_full) and contains_data(file_name_status_full):
                raise Exception("error file contains data: " + file_name_status_full, (
                    return_code + return_stdout + return_stderr))

        message = "Finished Import Process: " + bat_file + " for file: " + import_file
        print message

        import_successful = True

    return return_code + return_stdout + return_stderr

def export_dataloader(importer_directory, salesforce_type, interactivemode, displayalerts, location_local, client_type, client_subtype):
    
    """Export out of Salesforce using DataLoader"""

    from os.path import exists, join
    from subprocess import Popen, PIPE

    exporter_clientdirectory = importer_directory.replace("Importer", "Exporter")
    exporter_directory = exporter_clientdirectory
    if "\\Salesforce-Exporter\\" in exporter_directory:
        exporter_directory += "\\..\\..\\.."

    interactive_flag = ""
    if interactivemode:
        interactive_flag = "-interactivemode"
    bat_file = exporter_directory + "\\exporter.bat {} {}".format(salesforce_type, interactive_flag)

    return_code = ""
    return_stdout = ""
    return_stderr = ""

    if not exists(exporter_directory):
        print "Skip Export Process (export not detected)"
    else:
        message = "Starting Export Process: " + bat_file + "\n\nExport Process - can take up to a couple of minutes depending on your Internet connection..."
        print message
        return_stdout += message + "\n"
        export_process = Popen(bat_file, stdout=PIPE, stderr=PIPE)

        stdout, stderr = export_process.communicate()

        return_code += "\n\nexport_dataloader (returncode): " + str(export_process.returncode)
        return_stdout += "\n\nexport_dataloader (stdout):\n" + stdout
        return_stderr += "\n\nexport_dataloader (stderr):\n" + stderr

        if (export_process.returncode != 0
                or contains_error(return_stdout)
                or "We couldn't find the Java Runtime Environment (JRE)" in return_stdout):
            raise Exception("Invalid Return Code", return_code + return_stdout + return_stderr)

    #Check to extract the data from the content version if running in the cloud
    if not location_local:
        if not export_extractcontentexists(importer_directory, client_type, client_subtype):
            
            print "\nRunning in Cloud and no valid Import Instance so skip processing\n"
            global stop_processing
            stop_processing = True

    return return_code + return_stdout + return_stderr

def export_extractcontentexists(importer_directory, client_type, client_subtype):

    """Export - extract content exists - checks to see if running on cloud if there is any content scheduled for import"""
    import base64
    from csv import DictReader
    from os.path import join
    import sys
    import csv
    import os
    from subprocess import Popen, PIPE

    exporter_clientdirectory = join(importer_directory.replace("Importer", "Exporter"), "Export\\")
    linked_entity_ids = set()

    try:
        global emailattachments
        validImportInstance = False

        # Check for scheduled import instance
        with open(join(exporter_clientdirectory,'ImportInstanceExtract-Prod.csv'), 'r') as read_obj:
            csv_dict_reader = DictReader(read_obj)
            for row in csv_dict_reader:

                #Check for schedule related to current client
                if row['TYPE__C'] in client_subtype:

                    #Valid Import Instance but no files required so return without attempting to extract files
                    if row['EMAIL_ATTACH_LOGS__C'] == 'All Logs':
                        emailattachments = True
                    else:
                        emailattachments = False

                    validImportInstance = True
                    break

        # No valid import instance so return to kick out of process until there is a valid instance
        if not validImportInstance:
            return False

        # Attempt to extract file data
        with open(join(exporter_clientdirectory, 'ContentDocumentLinkExtract-Prod.csv'), 'r') as read_obj:
            csv_dict_reader = DictReader(read_obj)
            for row in csv_dict_reader:

                linked_entity_ids.add("'" + row['LINKEDENTITYID'] + "'")

    except Exception as ex:
        print "\nexport_extractcontent - Unexpected error:" + str(ex)

    if len(linked_entity_ids) <= 0:
        return True

    #run extract
    comma_list = ",".join(linked_entity_ids)
    p = Popen(['python', r'C:\repo\salesforce-files-download\download.py', '-o', exporter_clientdirectory, '-q', comma_list, '-t', client_type],
              stdout=PIPE,
              stderr=PIPE,
              cwd=r'C:\repo\salesforce-files-download')
    output = p.communicate()
    print output[0]

    return True

def export_odbc(importer_directory, salesforce_type, client_subtype, interactivemode, displayalerts):
    """Export out of Salesforce using DataLoader"""

    from os.path import exists
    from subprocess import Popen, PIPE

    exporter_directory = importer_directory.replace("Salesforce-Importer", "ODBC-Exporter")
    if "\\ODBC-Exporter\\" in exporter_directory:
        exporter_directory += "\\..\\..\\.."

    interactive_flag = ""
    if (interactivemode or displayalerts):
        interactive_flag = "-interactivemode"
    bat_file = exporter_directory + "\\exporter.bat {} {} {}".format(salesforce_type,
                                                                     client_subtype,
                                                                     interactive_flag)

    return_code = ""
    return_stdout = ""
    return_stderr = ""

    if not exists(exporter_directory):
        print "Skip ODBC Export Process (export not detected)"
    else:
        message = "Starting ODBC Export Process: " + bat_file
        print message
        return_stdout += message + "\n"
        export_process = Popen(bat_file, stdout=PIPE, stderr=PIPE)

        stdout, stderr = export_process.communicate()

        return_code += "\n\nexport_odbc (returncode): " + str(export_process.returncode)
        return_stdout += "\n\nexport_odbc (stdout):\n" + stdout
        return_stderr += "\n\nexport_odbc (stderr):\n" + stderr

        if (export_process.returncode != 0
                or contains_error(return_stdout)
                or "We couldn't find the Java Runtime Environment (JRE)" in return_stdout):
            raise Exception("Invalid Return Code", return_code + return_stdout + return_stderr)

    return return_code + return_stdout + return_stderr

def contains_error(text):
    """ Check for errors in text string """

    modified_text = text.lower().replace("0 errors", "success")

    errorFound = False

    if "error" in modified_text.lower():
        errorFound = True

    if "exception" in modified_text.lower():
        errorFound = True

    return errorFound

def send_email(client_emaillist, subject, file_path, emailattachments, log_path):
    """Send email via O365"""

    import base64
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import COMMASPACE, formatdate
    import os
    from os.path import basename
    from shutil import copy
    import smtplib
    import re

    message = "\n\nPreparing email results\n"
    print message

    send_to = client_emaillist.split(";")

    send_from = 'db.powerbi@501commons.org'
    server = "smtp.office365.com"

    if 'SERVER_EMAIL_USERNAME' in os.environ:
        send_from = os.environ.get('SERVER_EMAIL_USERNAME')

    if 'SERVER_EMAIL' in os.environ:
        server = os.environ.get('SERVER_EMAIL')

    #https://stackoverflow.com/questions/3362600/how-to-send-email-attachments

    msg = MIMEMultipart()

    msg['From'] = send_from
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    from os import listdir
    from os.path import isfile, join, exists

    msgbody = subject + "\n\n"
    if not emailattachments:
        msgbody += "Attachments disabled: Result files can be accessed on the import client.\n\n"

    # Send To Admin Only unless there is a csv file which means there was at least a load attempt and not a system failure
    sendTo_AdminOnly = True

    sendTo_AdminAddress = "sfconsulting@501commons.org"
    for sendToEmail in send_to:
        if re.search("501commons", sendToEmail, re.IGNORECASE):
            sendTo_AdminAddress = sendToEmail
            break

    if file_path:
        onlyfiles = [join(file_path, f) for f in listdir(file_path)
                    if isfile(join(file_path, f))]

        msgbody += "Log Directory: {}\n\n".format(log_path)

        for file_name in onlyfiles:

            if contains_data(file_name) and ".sent" not in file_name:

                msgbody += "\t{}, with {} rows\n".format(basename(file_name), file_linecount(file_name))

                if emailattachments or (contains_error(subject) and "log" in file_name.lower()) or contains_error(file_name.lower()):

                    if "csv" in file_name:
                        sendTo_AdminOnly = False

                    with open(file_name, "rb") as file_name_open:
                        part = MIMEApplication(
                            file_name_open.read(),
                            Name=basename(file_name)
                            )

                    # After the file is closed
                    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_name)
                    msg.attach(part)

                # Rename file so do not attached again
                sent_file = join(file_path, file_name)
                filename, file_extension = os.path.splitext(sent_file)
                sent_file = "{}.sent{}".format(filename, file_extension)

                if exists(sent_file):
                    os.remove(sent_file)

                os.rename(file_name, sent_file)

                #Save copy to log directory
                copy(sent_file, log_path)


    # Check if sending email only to the Admin
    if sendTo_AdminOnly:
        msg['To'] = sendTo_AdminAddress
    else:
        msg['To'] = COMMASPACE.join(send_to)

    import time
    msgbody += "\n\n501 Commons ETL Version: %s\n\n" % format(time.ctime(os.path.getmtime(join(file_path, '..\\..\\..\\importer.py'))))

    print msgbody
    msg.attach(MIMEText(msgbody))

    server = smtplib.SMTP(server, 587)
    server.starttls()

    server_password = 'unknown'
    if 'SERVER_EMAIL_PASSWORD' in os.environ:
        server_password = os.environ.get('SERVER_EMAIL_PASSWORD')

    if 'SERVER_EMAIL_PASSWORDOVERRIDE' in os.environ:
        server_password = os.environ.get('SERVER_EMAIL_PASSWORDOVERRIDE')

    server.login(send_from, base64.b64decode(server_password))
    text = msg.as_string()

    server.sendmail(send_from, send_to, text)
    server.quit()

    message = "\nSent email results\n"
    print message

def send_salesforce():
    """Send results to Salesforce to handle notifications"""
    #Future update to send to salesforce to handle notifications instead of send_email
    #https://developer.salesforce.com/blogs/developer-relations/2014/01/
    #python-and-the-force-com-rest-api-simple-simple-salesforce-example.html

if __name__ == "__main__":
    main()
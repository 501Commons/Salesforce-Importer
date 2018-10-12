"""import Module for Excel to Salesforce"""
global IS_WINDOWS

IS_WINDOWS = False
try:
    from win32api import STD_INPUT_HANDLE
    from win32console import GetStdHandle, KEY_EVENT, ENABLE_ECHO_INPUT, ENABLE_LINE_INPUT, ENABLE_PROCESSED_INPUT
    IS_WINDOWS = True
except ImportError as e:
    import sys
    import select
    import termios

class KeyboardHook():
    """Keyboard Hook Class"""
    def __enter__(self):
        global IS_WINDOWS
        if IS_WINDOWS:
            self.readHandle = GetStdHandle(STD_INPUT_HANDLE)
            self.readHandle.SetConsoleMode(ENABLE_LINE_INPUT|ENABLE_ECHO_INPUT|ENABLE_PROCESSED_INPUT)

            self.curEventLength = 0
            self.curKeysLength = 0

            self.capturedChars = []
        else:
            # Save the terminal settings
            self.fd = sys.stdin.fileno()
            self.new_term = termios.tcgetattr(self.fd)
            self.old_term = termios.tcgetattr(self.fd)

            # New terminal setting unbuffered
            self.new_term[3] = (self.new_term[3] & ~termios.ICANON & ~termios.ECHO)
            termios.tcsetattr(self.fd, termios.TCSAFLUSH, self.new_term)

        return self

    def __exit__(self, type, value, traceback):
        if IS_WINDOWS:
            pass
        else:
            termios.tcsetattr(self.fd, termios.TCSAFLUSH, self.old_term)

    def poll(self):
        """poll method to check for keyboard input"""
        if IS_WINDOWS:
            if not len(self.capturedChars) == 0:
                return True

            events_peek = self.readHandle.PeekConsoleInput(10000)

            if len(events_peek) == 0:
                return False

            if not len(events_peek) == self.curEventLength:
                for current_event in events_peek[self.curEventLength:]:
                    if current_event.EventType == KEY_EVENT:
                        if ord(current_event.Char) == 0 or not current_event.KeyDown:
                            pass
                        else:
                            curChar = str(current_event.Char)
                            self.capturedChars.append(curChar)
                self.curEventLength = len(events_peek)

            if not len(self.capturedChars) == 0:
                return True
            else:
                return False
        else:
            data_read, data_write, data_error = select.select([sys.stdin], [], [], 0)
            if not data_read == []:
                return True
            return None

def main():
    """Main entry point"""

    import sys
    import os
    from os import listdir, makedirs
    from os.path import exists, join

    print 'Wait Loop1'
    with KeyboardHook() as keyboard_hook:
        while True:
            keyboard_input = keyboard_hook.poll()
            if not keyboard_input is None:
                print "\nUser interrupted wait cycle\n"
                break

    print 'Wait Loop2'
    with KeyboardHook() as keyboard_hook:
        while True:
            keyboard_input = keyboard_hook.poll()
            if not keyboard_input is None:
                print "\nUser interrupted wait cycle\n"
                break

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

    print ("Incoming required paramters: " +
           "salesforce_type: {} client_type: {} client_subtype: {} client_emaillist: {}\n"
           .format(salesforce_type, client_type, client_subtype, client_emaillist))

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

    noexportodbc = False
    if '-noexportodbc' in sys.argv:
        noexportodbc = True

    noexportsf = False
    if '-noexportsf' in sys.argv:
        noexportsf = True

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

    #Clear out log directory
    importer_log_directory = join(importer_root, "..\\Status\\")
    if not exists(importer_log_directory):
        makedirs(importer_log_directory)

    importer_log_directory = join(importer_log_directory, client_subtype)
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

    # Insert Data
    status_import = ""
    if not norefresh and "Invalid Return Code" not in status_export:
        for insert_run in range(0, insert_attempts):

            print "\n\nImporter - Insert Data Process (run: %d)\n\n" % (insert_run)

            status_import = process_data(importer_directory, salesforce_type, client_type,
                                         client_subtype, False, wait_time,
                                         noexportsf,
                                         interactivemode,
                                         displayalerts,
                                         skipexcelrefresh)

            # Insert files are empty so continue to update process
            if "import_dataloader (returncode)" not in status_import:
                break

    # Update Data
    if not noupdate and not contains_error(status_import):
        print "\n\nImporter - Update Data Process\n\n"
        status_import = process_data(importer_directory, salesforce_type, client_type,
                                     client_subtype, True, wait_time,
                                     noexportsf, interactivemode, displayalerts, skipexcelrefresh)

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

    # Send email results
    results = "Success"
    if contains_error(status_import) or contains_error(status_export):
        results = "Error"
    subject = "{}-{} Salesforce Importer Results - {}".format(client_type, client_subtype, results)
    send_email(client_emaillist, subject, file_path, emailattachments, importer_log_directory)

    print "\nImporter process completed\n"

def process_data(importer_directory, salesforce_type, client_type,
                 client_subtype, update_mode, wait_time,
                 noexportsf, interactivemode, displayalerts, skipexcelrefresh):
    """Process Data based on data_mode"""

    #Create log file for import status and reports
    from os import makedirs
    from os.path import exists, join
    file_path = importer_directory + "\\Status"
    if not exists(file_path):
        makedirs(file_path)

    data_mode = "Insert"
    if update_mode:
        data_mode = "Update"

    output_log = "Process Data (" + data_mode + ")\n\n"
    status_process_data = ""

    # Export data from Salesforce

    try:
        if not noexportsf:
            status_process_data = export_dataloader(importer_directory,
                                                    salesforce_type, interactivemode, displayalerts)
        else:
            status_process_data = "Skipping export from Salesforce"
    except Exception as ex:
        output_log += "\n\nexport_dataloader - Unexpected export error:" + str(ex)
        status_process_data = "Error detected - Exception"
    else:
        output_log += "\n\nExport\n" + status_process_data

    # Export data from Excel

    try:
        if (not skipexcelrefresh and not contains_error(status_process_data)
                and not contains_error(output_log.lower())):
            status_process_data = refresh_and_export(importer_directory,
                                                     salesforce_type, client_type,
                                                     client_subtype, update_mode,
                                                     wait_time, interactivemode, displayalerts)
        else:
            status_process_data = "Skipping refresh and export from Excel"
    except Exception as ex:
        output_log += "\n\nrefresh_and_export - Unexpected export error:" + str(ex)
        status_process_data = "Error detected - Exception"
    else:
        output_log += "\n\nExport\n" + status_process_data

    # Import Data into Salesforce

    try:
        if not contains_error(status_process_data) and not contains_error(output_log):
            status_process_data = import_dataloader(importer_directory,
                                                    client_type, salesforce_type,
                                                    data_mode)
        else:
            status_process_data = "Error detected so skipped"
    except Exception as ex:
        output_log += "\n\nUnexpected import error:" + str(ex)
        status_process_data = "Error detected - Exception"
    else:
        output_log += "\n\nImport\n" + status_process_data

    import datetime
    date_tag = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    with open(join(file_path, "Salesforce-Importer-Log-{}-{}.txt".format(data_mode, date_tag)),
              "w") as text_file:
        text_file.write(output_log)

    return status_process_data

def refresh_and_export(importer_directory, salesforce_type,
                       client_type, client_subtype, update_mode,
                       wait_time, interactivemode, displayalerts):
    """Refresh Excel connections"""

    import os
    import os.path
    import time
    import win32com.client as win32

    try:
        refresh_status = "refresh_and_export\n"
        excel_connection = win32.gencache.EnsureDispatch("Excel.Application")
        excel_file_path = importer_directory + "\\"
        workbooks = excel_connection.Workbooks
        workbook = workbooks.Open((
            excel_file_path + client_type + "-" + client_subtype + "_" + salesforce_type + ".xlsx"))

        excel_connection.Visible = interactivemode
        excel_connection.DisplayAlerts = displayalerts

        #for connection in workbook.Connections:
            #print connection.name
            # BackgroundQuery does not work so have to do manually in Excel for each Connection
            #connection.BackgroundQuery = False

        # RefreshAll is Synchronous iif
        #   1) Enable background refresh disabled/unchecked in xlsx for all Connections
        #   2) Include in Refresh All enabled/checked in xlsx for all Connections
        #   To verify: Open xlsx Data > Connections > Properties for each to verify
        message = "\nRefreshing all connections..."
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
        message = ("Pausing " + str(wait_time) +
                   " seconds to give Excel time to complete data queries..." +
                   "\n\t\t***if Excel Refresh complete then press any key to exit wait cycle")
        print message
        refresh_status += message + "\n"

        with KeyboardHook() as keyboard_hook:
            while wait_time > 0:
                if wait_time > 30:
                    time.sleep(30)

                    wait_time -= 30
                    message = ("\t" + str(wait_time) +
                               " seconds remaining for Excel to complete data queries..." +
                               "\n\t\t***if Excel Refresh complete then press any key to exit wait cycle")
                    print message
                    refresh_status += message + "\n"

                else:
                    time.sleep(wait_time)
                    wait_time = 0
                    break

                keyboard_input = keyboard_hook.poll()
                if not keyboard_input is None:
                    print "\nUser interrupted wait cycle\n"
                    break

        message = "Refreshing all connections...Completed"
        print message
        refresh_status += message + "\n"

        if not os.path.exists(excel_file_path + "Import\\"):
            os.makedirs(excel_file_path + "Import\\")

        for sheet in workbook.Sheets:
            # Only export update, insert, or report sheets
            sheet_name_lower = sheet.Name.lower()
            if ("update" not in sheet_name_lower
                    and "insert" not in sheet_name_lower
                    and "report" not in sheet_name_lower):
                continue

            excel_connection.Sheets(sheet.Name).Select()
            sheet_file = excel_file_path + "Import\\" + sheet.Name + ".csv"

            message = "Exporting csv for sheet: " + sheet_file
            print message
            refresh_status += message + "\n"

            # Save report to Status to get attached to email
            if "report" in sheet.Name.lower():
                sheet_file = excel_file_path + "Status\\" + sheet.Name + ".csv"

            # Check for existing file
            if os.path.isfile(sheet_file):
                os.remove(sheet_file)

            workbook.SaveAs(sheet_file, 6)

            # Update check to make sure insert sheet is empty
            if update_mode and "insert" in sheet.Name.lower() and contains_data(sheet_file):
                raise Exception("Update Error", (
                    "Insert sheet contains data and should be empty during update process: " +
                    sheet_file))

    except Exception as ex:
        refresh_status += "Unexpected error:" + str(ex)
        raise Exception("Export Error", refresh_status)

    finally:
        workbook.Close(False)
        # Marshal.ReleaseComObject(workbooks)
        # Marshal.ReleaseComObject(workbook)
        # Marshal.ReleaseComObject(excel_connection)
        excel_connection.Quit()

    return refresh_status

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

def import_dataloader(importer_directory, client_type, salesforce_type, data_mode):
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

    for file_name in listdir(bat_path):
        if not data_mode in file_name or ".sdl" not in file_name:
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

    return return_code + return_stdout + return_stderr

def export_dataloader(importer_directory, salesforce_type, interactivemode, displayalerts):
    """Export out of Salesforce using DataLoader"""

    from os.path import exists
    from subprocess import Popen, PIPE

    exporter_directory = importer_directory.replace("Importer", "Exporter")
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
        message = "Starting Export Process: " + bat_file
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

    return return_code + return_stdout + return_stderr

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

    if "error" in modified_text.lower():
        return True

    if "exception" in modified_text.lower():
        return True

    return False

def send_email(client_emaillist, subject, file_path, emailattachments, log_path):
    """Send email via O365"""

    message = "\n\nPreparing email results\n"
    print message

    send_to = client_emaillist.split(";")
    send_from = 'db.powerbi@501commons.org'
    server = "smtp.office365.com"

    #https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
    import base64
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import COMMASPACE, formatdate
    import os
    from os.path import basename
    from shutil import copy
    import smtplib

    msg = MIMEMultipart()

    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    from os import listdir
    from os.path import isfile, join, exists

    onlyfiles = [join(file_path, f) for f in listdir(file_path)
                 if isfile(join(file_path, f))]

    msgbody = subject + "\n\n"
    if not emailattachments:
        msgbody += "Attachments disabled: Result files can be accessed on the import client.\n\n"

    msgbody += "Log Directory: {}\n\n".format(log_path)

    for file_name in onlyfiles:

        if contains_data(file_name) and ".sent" not in file_name:

            msgbody += "\t{}, with {} rows\n".format(basename(file_name), file_linecount(file_name))

            if emailattachments or (contains_error(subject) and "log" in file_name.lower()):
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

    print msgbody
    msg.attach(MIMEText(msgbody))

    server = smtplib.SMTP(server, 587)
    server.starttls()
    server_password = os.environ['SERVER_EMAIL_PASSWORD']
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
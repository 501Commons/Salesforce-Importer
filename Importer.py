"""import Module for Excel to Salesforce"""
def main():
    """Main entry point"""

    import sys
    from os.path import join, dirname, realpath

    salesforce_type = str(sys.argv[1])
    client_type = str(sys.argv[2])
    client_subtype = str(sys.argv[3])
    client_emaillist = str(sys.argv[4])

    importer_directory = join(dirname(realpath(__file__)), "Clients\\" + client_type)

    # Insert Data
    process_data(importer_directory, salesforce_type, client_type,
                 client_subtype, False, client_emaillist)

    # Update Data
    process_data(importer_directory, salesforce_type, client_type,
                 client_subtype, True, client_emaillist)

def process_data(importer_directory, salesforce_type, client_type,
                 client_subtype, update_mode, client_emaillist):
    """Process Data based on data_mode"""

    from os import makedirs
    from os.path import exists

    data_mode = "Insert"
    if update_mode:
        data_mode = "Update"

    client_email = "db." + client_type + "@501commons.org"
    sendto = [client_email, 'daveb@501commons.org']
    sendto += client_emaillist.split(";")

    user = 'db.powerbi@501commons.org'
    smtpsrv = "smtp.office365.com"
    subject = "Process Data (" + data_mode + ") Results -"
    file_path = importer_directory + "\\Status"
    if not exists(file_path):
        makedirs(file_path)

    body = "Process Data (" + data_mode + ")\n\n"

    # Export data from Excel
    try:
        status_export = refresh_and_export(importer_directory, salesforce_type, client_type,
                                           client_subtype, update_mode)
    except Exception as ex:
        subject += " Error Export"
        body += "Unexpected export error:" + str(ex)
    else:
        body += "Export\n" + status_export

    # Import data into Salesforce
    try:
        if not "Error" in subject:
            status_import = import_dataloader(importer_directory, client_type, data_mode)
        else:
            status_import = "Error detected so skipped"
    except Exception as ex:
        subject += " Error Import"
        body += "\n\nUnexpected import error:" + str(ex)
    else:
        body += "\n\nImport\n" + status_import

    if not "Error" in subject:
        subject += " Successful"

    # Send email results
    send_email(user, sendto, subject, body, file_path, smtpsrv)

def refresh_and_export(importer_directory, salesforce_type, client_type, client_subtype, update_mode):
    """Refresh Excel connections"""

    #import datetime
    import os
    import os.path
    import win32com.client

    try:
        refresh_status = "refresh_and_export\n"
        excel_connection = win32com.client.DispatchEx("Excel.Application")
        excel_file_path = importer_directory + "\\"
        workbook = excel_connection.workbooks.open((
            excel_file_path + client_type + "-" + client_subtype + "_" + salesforce_type + ".xlsx"))

        # Uncomment if you want to see the Excel file opened
        #excel_connection.Visible = True

        #for connection in workbook.Connections:
            #print connection.name
            # BackgroundQuery does not work so have to do manually in Excel for each Connection
            #connection.BackgroundQuery = False

        # RefreshAll is Syncronous iif
        #   1) Enable background refresh disabled/unchecked in xlsx for all Connections
        #   2) Include in Refresh All enabled/checked in xlsx for all Connections
        #   To verify: Open xlsx Data > Connections > Properties for each to verify
        message = "Refreshing all connections..."
        print message
        refresh_status += message + "\n"

        #workbook.RefreshAll()
        message = "Refreshing all connections...Completed"
        print message
        refresh_status += message + "\n"

        if not os.path.exists(excel_file_path + "Import\\"):
            os.makedirs(excel_file_path + "Import\\")

        #date_tag = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        for sheet in workbook.Sheets:
            # Only export update or insert sheets
            if "update" not in sheet.name.lower() and "insert" not in sheet.name.lower():
                continue

            message = "Exporting csv for sheet: " + sheet.name
            print message
            refresh_status += message + "\n"

            excel_connection.Sheets(sheet.name).Select()
            sheet_file = excel_file_path + "Import\\" + sheet.name + ".csv"

            # Check for existing file
            if os.path.isfile(sheet_file):
                os.remove(sheet_file)

            workbook.SaveAs(sheet_file, 6)

            # Update check to make sure insert sheet is empty
            if update_mode and "insert" in sheet.name.lower() and contains_data(sheet_file):
                raise Exception("Update Error", (
                    "Insert sheet contains data and should be empty during update process: " +
                    sheet_file))

    except Exception as ex:
        refresh_status += "Unexpected error:" + str(ex)
        raise Exception("Export Error", refresh_status)

    finally:
        workbook.Close(True)
        excel_connection.Quit()

    return refresh_status

def contains_data(file_name):
    """Check if file contains data after header"""

    line_index = 1
    with open(file_name) as file_open:
        for line in file_open:
            # Check if line empty or if line missing data
            if line_index == 2 and line != "" and not ",," in line and not '"",""' in line:
                return True
            elif line_index > 2:
                return True

            line_index += 1

    return False

def import_dataloader(importer_directory, client_type, data_mode):
    """Import into Salesforce using DataLoader"""

    import os
    from os import listdir
    from os.path import join
    from subprocess import Popen, PIPE

    bat_path = importer_directory
    import_path = importer_directory + "\\Import"

    return_code = ""
    return_stdout = ""
    return_stderr = ""

    for file_name in listdir(bat_path):
        if not data_mode in file_name:
            continue

        # Check if associated csv has any data
        import_file = join(import_path, os.path.splitext(file_name)[0] + ".csv")
        if not os.path.exists(import_file) or not contains_data(import_file):
            continue

        bat_file = join(bat_path, file_name)
        import_process = Popen(bat_file, stdout=PIPE, stderr=PIPE)
        stdout, stderr = import_process.communicate()

        return_code += "import_dataloader (returncode): " + str(import_process.returncode)
        return_stdout += "\n\nimport_dataloader (stdout):\n" + stdout
        return_stderr += "\n\nimport_dataloader (stderr):\n" + stderr

        if import_process.returncode != 0:
            raise Exception("Invalid Return Code", return_code + return_stdout + return_stderr)

        status_path = ("C:\\repo\\Salesforce-Importer\\Clients\\" +
                       client_type + "\\status")
        for file_name_status in listdir(status_path):
            file_name_status_full = join(status_path, file_name_status)
            if "error" in file_name_status_full and contains_data(file_name_status_full):
                raise Exception("error file contains data: " + file_name_status_full, (
                    return_code + return_stdout + return_stderr))

    return return_code + return_stdout + return_stderr

def send_email(send_from, send_to, subject, text, file_path, server):
    """Send email via O365"""

    #https://stackoverflow.com/questions/3362600/how-to-send-email-attachments
    import base64
    import os
    import smtplib
    from os.path import basename, exists
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.utils import COMMASPACE, formatdate

    msg = MIMEMultipart()

    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(text))

    from os import listdir, remove
    from os.path import isfile, join
    onlyfiles = [join(file_path, f) for f in listdir(file_path)
                 if isfile(join(file_path, f))]

    for file_name in onlyfiles:
        with open(file_name, "rb") as file_name_open:
            part = MIMEApplication(
                file_name_open.read(),
                Name=basename(file_name)
                )
        # After the file is closed
        part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file_name)
        msg.attach(part)

    server = smtplib.SMTP(server, 587)
    server.starttls()
    server_password = os.environ['SERVER_EMAIL_PASSWORD']
    server.login(send_from, base64.b64decode(server_password))
    text = msg.as_string()
    #server.sendmail(send_from, send_to, text)
    server.quit()

    # Delete all status files
    for file_name in onlyfiles:
        try:
            remove(file_name)
        except:
            continue

    # Delete all import files
    import_path = join(file_path, "..\\Import")
    if exists(import_path):
        for file_name in listdir(import_path):
            try:
                remove(join(import_path, file_name))
            except:
                continue

def send_salesforce():
    """Send results to Salesforce to handle notifications"""
    #Future update to send to salesforce to handle notifications instead of send_email
    #https://developer.salesforce.com/blogs/developer-relations/2014/01/python-and-the-force-com-rest-api-simple-simple-salesforce-example.html

if __name__ == "__main__":
    main()

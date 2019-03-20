

Data Loader
https://help.salesforce.com/articleView?id=data_loader.htm

Setup Instructions

1) Install Zulu & Salesforce Data Loader (Instructions -> https://help.salesforce.com/articleView?id=loader_install_windows.htm)

2) Install the latest Java Platform (JDK) -> https://www.oracle.com/technetwork/java/javase/downloads/index.html 

3) Install Git for Windows -> http://gitforwindows.org
    NOTE: Don't need to have an account just need the application installed

4) Install Python 2.7.14 https://www.python.org/downloads/ 

5) **501 Admin** will provider zip file for custom settings ([Client].zip).  Extract zip into C:\repo\Salesforce-Exporter-Private\Clients\[Client].

Example: C:\repo\Salesforce-Exporter-Private\Clients\XYZ where XYZ are the ClientInitials should contain an Exporter.bat file and a DataLoader directory.

6) Edit c:\repo\Exporter-Private\Exporter.bat
    Check & Verify the following values - update accordingly
    * EMAIL_LIST - include client emails
    * IMPORT_DIRECTORY - Location of incoming data files (e.g., DropBox, OneDrive)
    * JAVA_HOME - Verify directory is valid or change to correct xxx version number based on installed C:\Program Files\Java\jre1.8.0_xxx\bin

7) **501 Admin** will provider zip file for custom settings ([Client].zip).  Extract zip into C:\repo\Salesforce-Importer-Private\Clients\[Client].

Example: C:\repo\Salesforce-Importer-Private\Clients\XYZ where XYZ are the ClientInitials should contain an importer.bat file and a DataLoader directory.

7) Edit c:\repo\Importer-Private\importer.bat
    Check & Verify the following values - update accordingly
    * EMAIL_LIST - include client emails
    * IMPORT_DIRECTORY - Location of incoming data files (e.g., DropBox, OneDrive)
    * JAVA_HOME - Verify directory is valid or change to correct xxx version number based on installed C:\Program Files\Java\jre1.8.0_xxx\bin

Running Import

Run c:\repo\Importer-Private\importer.bat to start the importer.  You can run
    - importer.bat manually
    - schedule with Task Scheduler (be sure to set working directory to the importer.bat directory)

Troubleshooting

1) Excel gives an error when trying to authenticate with Salesforce
Resolution: Enable Excel to use TLS 1.2
https://social.technet.microsoft.com/Forums/en-US/92811d44-1165-4da2-96e7-20dc99bdf718/can-power-query-be-updated-to-use-tls-version-12?forum=powerquery

2) Importer process keeps popping up Privacy Levels dialog.
Resolution: Checking ignore doesn't always stop further prompting so set to Public if you keep getting prompted.  Another option is open the Excel file in C:\repo\Salesforce-Importer-Private\Clients\[Client]\Salesforce-Importer\Clients\[Client] and set the Privacy levels then save the Excel file.

3) Salesforce Data Loader can't install Admin version to C:\Program Files (x86)\salesforce.com
Resolution: You can install on another machine where you are an administrator and then just copy the salesforce.com directory to C:\Program Files (x86) to your target machine.

4) Running importer and getting an error in the console window that says, "Unlink of file '[excelfile]].xlsx' failed. Should I try again? (y/n)"
Resolution: Importer did not properly close the previous Excel session(s) - restarting your computer will solve the problem.

5) Importer Email says - Error Import and within the email there is a 'java.lang.RuntimeException: java.lang.NullPointerException'
Resolution: This is probably related to empty columns in the generated CSV file.  Open your Excel file and save out each sheet then edit the CSV files in Notepad.  If you see a bunch of ,,,,,,,,,, in the header columns then that is the issue.  To fix open Excel and select all the empty columns after your last data column and delete the columns.  The other method is delete the sheet and then right click on your Data source and Load to... a new sheet.
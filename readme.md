
Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions

1) Verify Microsoft Excel 2016 is installed and working 

2) Install Salesforce Data Loader (Instructions -> https://help.salesforce.com/articleView?id=000239784&type=1)

NOTE:
  a) Make sure your current user has Administrator access on the machine
  b) During Installation on the 'Install for admins only?' screen when prompted for 'Do you have administrator rights on this machine?' select Yes

3) Run Salesforce Data Loader to verify installation.  If you need Java installed then you will be prompted to install Java and follow the process to install Java. After Java installed run Data Loader to verify installation. 

4) Install Git for Windows http://gitforwindows.org
    NOTE: Don't need to have an account just need the application installed

5) Install Python 2.7.14 https://www.python.org/downloads/ 

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
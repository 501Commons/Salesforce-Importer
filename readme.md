
Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions

1) Install Git for Windows http://gitforwindows.org

2) Install Java JRE 1.8 http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html 

3) Install Python 2.7.14 https://www.python.org/downloads/ 

4) Install Salesforce Data Loader https://help.salesforce.com/articleView?id=000239784&type=1

5) **501 Admin** will provider zip file for custom settings ([Client].zip).  Extract zip into C:\repo\Salesforce-Importer-Private\Clients\[Client].

Example: C:\repo\Salesforce-Importer-Private\Clients\XYZ where XYZ are the ClientInitials should contain an importer.bat file and a DataLoader directory.

6) Edit c:\repo\Importer-Private\importer.bat
    Check & Verify the following values - update accordingly
    * EMAIL_LIST - include client emails
    * IMPORT_DIRECTORY - Location of incoming data files (e.g., DropBox, OneDrive)
    * JAVA_HOME - Verify directory is valid or change to correct xxx version number based on installed C:\Program Files\Java\jre1.8.0_xxx\bin

Running Import

Run c:\repo\Importer-Private\importer.bat to start the importer.  You can run
    - importer.bat manually
    - schedule with Task Scheduler (be sure to set working directory to the importer.bat directory)
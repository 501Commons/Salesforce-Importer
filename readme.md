
Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions

1) Install Git for Windows http://gitforwindows.org

2) Install Java JRE 1.8 http://www.oracle.com/technetwork/java/javase/downloads/jre8-downloads-2133155.html 

3) Install Python 2.7.14 https://www.python.org/downloads/ 

4) Install Salesforce Data Loader https://help.salesforce.com/articleView?id=000239784&type=1

5) **501 Admin** will provider zip file for custom settings.  Extract zip ideally to a Cloud Storage location w/ 501 Commons Access otherwise to c:\repo\Importer-Private.

6) Edit c:\repo\Importer-Private\importer.bat
    Check & Verify the following values - update accordingly
    * EMAIL_LIST - include client emails
    * IMPORT_DIRECTORY - change to client Import Location
    * JAVA_HOME
    * PYTHON_HOME

Running Import

Run c:\repo\Importer-Private\importer.bat to start the importer.  You can run
    - importer.bat manually
    - schedule with Task Scheduler (be sure to set working directory to the importer.bat directory)

Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions

1) Install Git for Windows http://gitforwindows.org

2) Install Java JRE 1.8

3) Install Python 2.7.14 https://www.python.org/downloads/ 

4) **501 Admin** will provider zip file for custom settings.  Extract zip to c:\repo\Importer-Private

Running Import

Run c:\repo\importer.bat to start the importer.  You can run
    - importer.bat manually
    - schedule with Task Scheduler (be sure to set working directory to the importer.bat directory)
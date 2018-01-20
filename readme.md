
Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions

1) Install Git for Windows http://gitforwindows.org

2) Install Java JRE 1.8

3) Set JAVA_HOME Environment Variable
Example: JAVA_HOME = C:\Program Files\Java\jre1.8.0_151\bin

4) Open 'Git CMD' and type: pip install pypiwin32

5) **501 Admin** will provider zip file for custom settings.  Extract zip to c:\repo\Importer-Private

Running Import

Run c:\repo\importer.bat to start the importer.  You can run
    - importer.bat manually
    - schedule with Task Scheduler (be sure to set working directory to the importer.bat directory)
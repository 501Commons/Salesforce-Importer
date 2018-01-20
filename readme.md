
Data Loader Guide
http://resources.docs.salesforce.com/210/17/en-us/sfdc/pdf/salesforce_data_loader.pdf

Data Loader Quick Reference
http://www.developerforce.com/media/Cheatsheet_Setting_Up_Automated_Data_Loader_9_0.pdf

Setup Instructions
1) Install Java JRE 1.8

2) Set JAVA_HOME Environment Variable
Example: JAVA_HOME = C:\Program Files\Java\jre1.8.0_151\bin

3) pip install pypiwin32

4) Install Git for Windows http://gitforwindows.org

5) run 'Git CMD'

6) mkdir c:\repo

7) cd c:\repo

8) git clone https://github.com/501commons/Salesforce-Importer.git 

9) **501 Admin** will provider zip file for custom settings.  Extract zip to c:\repo\Salesforce-Importer\Clients\[Client Type]

Running import

run Clients\[Client Type]\importer.bat to start the importer.  You can run importer.bat manually or schedule with Task Scheduler
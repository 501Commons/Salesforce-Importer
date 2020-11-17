@REM Create Shift Coverage Export
cd C:\repo\Salesforce-Exporter-Private\Clients\IISC\
call exporter.bat Prod db.iisc@501commons.org

@REM Run SC Importer
cd C:\repo\Salesforce-Importer-Private\Clients\IISC\
call importer.bat Compass-Insert Prod Cloud

@REM Run SC Importer
cd C:\repo\Salesforce-Importer-Private\Clients\IISC\
call importer.bat Compass-Update1of2 Prod Cloud
call importer.bat Compass-Update2of2 Prod Cloud

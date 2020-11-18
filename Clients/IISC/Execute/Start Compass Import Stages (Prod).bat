@REM Run SC Importer Insert
call C:\repo\Salesforce-Importer-Private\Clients\IISC\importer.bat Compass-Insert Prod Cloud

@REM Run SC Importer Update
call C:\repo\Salesforce-Importer-Private\Clients\IISC\importer.bat Compass-Update1of2 Prod Cloud
call C:\repo\Salesforce-Importer-Private\Clients\IISC\importer.bat Compass-Update2of2 Prod Cloud

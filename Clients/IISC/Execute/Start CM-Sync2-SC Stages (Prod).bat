
@REM Create Class Management Export
set EMAIL_LIST=db.iisc@501commons.org;pauli@imaginewa.org
call C:\repo\Salesforce-Exporter-Private\Clients\II\exporter.bat Prod %EMAIL_LIST%

@REM Run SC Importer Insert
call C:\repo\Salesforce-Importer-Private\Clients\IISC\importer.bat SYNC-Insert Prod Cloud

@REM Run SC Importer Update
call C:\repo\Salesforce-Exporter-Private\Clients\IISC\importer.bat SYNC-Update1of2 Prod Cloud
call C:\repo\Salesforce-Importer-Private\Clients\IISC\importer.bat SYNC-Update2of2 Prod Cloud

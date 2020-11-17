
@REM Create Class Management Export
set EMAIL_LIST=db.iisc@501commons.org;pauli@imaginewa.org
cd C:\repo\Salesforce-Exporter-Private\Clients\II\
call exporter.bat Prod db.iisc@501commons.org

@REM Create Shift Coverage Export
cd C:\repo\Salesforce-Exporter-Private\Clients\IISC\
call exporter.bat Prod db.iisc@501commons.org

@REM Run SC Importer
cd C:\repo\Salesforce-Importer-Private\Clients\IISC\
call importer.bat SYNC-Insert Prod Cloud

@REM Run SC Importer
cd C:\repo\Salesforce-Importer-Private\Clients\IISC\
call importer.bat SYNC-Update1of2 Prod Cloud
call importer.bat SYNC-Update2of2 Prod Cloud

echo incoming %1 %2 %3
set DATALOADER_BIN=C:\repo\Salesforce-Importer-Private\Clients\%2\Salesforce-Importer\dependencies\salesforce.com\dataloader\v53.0.2\bin
cd %DATALOADER_BIN%
call %DATALOADER_BIN%\process.bat C:\repo\Salesforce-Importer-Private\Clients\%2\Salesforce-Importer\Clients\%2\DataLoader\ %3_%1

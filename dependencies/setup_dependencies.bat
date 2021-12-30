IF "%JAVA_HOME%" == "" (
    set JAVA_HOME=C:\Program Files\Zulu\zulu-17
    echo exporter: setting JAVA_HOME to %JAVA_HOME%
) ELSE (
    IF NOT EXIST "%JAVA_HOME%" (
        set JAVA_HOME=C:\Program Files\Zulu\zulu-17
        echo exporter: setting JAVA_HOME to %JAVA_HOME%
    )
)

set PYTHON_HOME=C:\Python27
set PATH=C:\Python27;C:\Python27\Scripts;%PATH%
echo exporter: setting PYTHON_HOME to %PYTHON_HOME%

IF "%SF_DATALOADER%" == "" (
    set SF_DATALOADER=C:\repo\Salesforce-Importer-Private\Clients\%CLIENT_TYPE%\Salesforce-Importer\dependencies\salesforce.com\dataloader\v53.0.2
    echo exporter: setting SF_DATALOADER to %SF_DATALOADER%
) ELSE (
    IF NOT EXIST "%SF_DATALOADER%" (
        set SF_DATALOADER=C:\repo\Salesforce-Importer-Private\Clients\%CLIENT_TYPE%\Salesforce-Importer\dependencies\salesforce.com\dataloader\v53.0.2
        echo exporter: setting SF_DATALOADER to %SF_DATALOADER%
    )
)
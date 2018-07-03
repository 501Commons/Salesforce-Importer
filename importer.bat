@echo off

set SF_DATALOADER=C:\Program Files (x86)\salesforce.com\Data Loader
IF NOT EXIST "%SF_DATALOADER%" (
    ECHO Error: "%SF_DATALOADER%" does not exist
    cscript importer.vbs "Salesforce Data Loader Required - a browser will open with the install info."
    call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
    goto scriptexit
)

REM Java should get installed as part of Data Loader
IF NOT EXIST "%JAVA_HOME%" (
    ECHO Error: "%JAVA_HOME%" does not exist
    cscript importer.vbs "Java not found which is installed via Salesforce Data Loader - Reinstall Data Loader and follow the instructions for Java setup.  A browser will open with the install info."
    call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
    goto scriptexit
)

IF NOT EXIST "%PYTHON_HOME%" (
    ECHO Error Python 2.7.14 Required: "%PYTHON_HOME%" does not exist
    cscript importer.vbs "Python 2.7.14 Required - a browser will open with the install info."
    call explorer "https://www.python.org/downloads/"
    goto scriptexit
)

IF "%IMPORT_DIRECTORY%" == "" (
    goto skip_import_directory_check
)

IF NOT EXIST "%IMPORT_DIRECTORY%" (
    cscript importer.vbs "Error Import Directory does not exist: %IMPORT_DIRECTORY%"
    goto scriptexit
)

:skip_import_directory_check

set PATH=%PATH%;%JAVA_HOME%;%PYTHON_HOME%;%PYTHON_HOME%\Scripts

cd "%PYTHON_HOME%"\Scripts
pip install pypiwin32

xcopy "%IMPORT_DIRECTORY%" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" /s /y /i
copy /Y "%IMPORTER_PRIVATE_DIR%\DataLoader\key.txt" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\DataLoader\key.txt"

echo ***************
IF "%IMPORT_ENVIRONMENT%" == "Sandbox" (
    echo *****Sandbox Data Import Automation
    python "%IMPORTER_DIRECTORY%\..\importer_sandbox.py" %IMPORT_ENVIRONMENT% %CLIENT_TYPE% %IMPORT_MODE% %EMAIL_LIST% %IMPORT_WAITTIME% %IMPORT_NOUPDATE% %IMPORT_NOEXPORTODBC% %IMPORT_NOEXPORTSF% %IMPORT_INSERTATTEMPTS%
) else (
    echo *****Production Data Import Automation
    python "%IMPORTER_DIRECTORY%\..\importer.py" %IMPORT_ENVIRONMENT% %CLIENT_TYPE% %IMPORT_MODE% %EMAIL_LIST% %IMPORT_WAITTIME% %IMPORT_NOUPDATE% %IMPORT_NOEXPORTODBC% %IMPORT_NOEXPORTSF% %IMPORT_INSERTATTEMPTS%
)
echo ***************

cd %IMPORTER_PRIVATE_DIR%

:scriptexit
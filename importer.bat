@echo off
SETLOCAL ENABLEDELAYEDEXPANSION

set SF_DATALOADER=C:\Program Files (x86)\salesforce.com\Data Loader
IF NOT EXIST "%SF_DATALOADER%" (
    ECHO Error: "%SF_DATALOADER%" does not exist
    cscript importer.vbs "Salesforce Data Loader Required - a browser will open with the install info."
    call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
    goto scriptexit
)

REM Java should get installed as part of Data Loader
IF EXIST "%JAVA_HOME%" goto java_exists

REM Attempt to find Java Version
call java.exe -version >temp.txt 2>&1
set /p JAVA_TEST=<temp.txt
FOR /F delims^=^"^ tokens^=2 %%G IN ("%JAVA_TEST%") DO set JAVA_VERSION=%%G
set JAVA_HOME_CHECK="C:\Program Files (x86)\Java\jre%JAVA_VERSION%"
IF EXIST "%JAVA_HOME_CHECK%" (
    set JAVA_HOME=%JAVA_HOME_CHECK%
    goto java_exists
) 

ECHO Error: "%JAVA_HOME%" does not exist
cscript importer.vbs "Java not found which is installed via Salesforce Data Loader - Reinstall Data Loader and follow the instructions for Java setup.  A browser will open with the install info."
call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
goto scriptexit

:java_exists

IF EXIST "%PYTHON_HOME%" goto python_exists

ECHO Error Python 2.7.14 Required: "%PYTHON_HOME%" does not exist
cscript importer.vbs "Python 2.7.14 Required - a browser will open with the install info."
call explorer "https://www.python.org/downloads/"
goto scriptexit

:python_exists

REM Using ! instead of % in case using special chars like parenthesis in path
IF "!IMPORT_DIRECTORY!" == "" (
    goto skip_import_directory_check
)

IF NOT EXIST "!IMPORT_DIRECTORY!" (
    cscript importer.vbs "Error Import Directory does not exist: !IMPORT_DIRECTORY!"
    goto scriptexit
)

REM Backward Compatibility: Try with and wihout quotes in case they are already included
xcopy "%IMPORT_DIRECTORY%" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" /s /y /i
if NOT EXIST "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" (
    xcopy %IMPORT_DIRECTORY% "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" /s /y /i
)

:skip_import_directory_check

set PATH=%PATH%;%JAVA_HOME%;%PYTHON_HOME%;%PYTHON_HOME%\Scripts

cd "%PYTHON_HOME%"\Scripts
python -m pip install --upgrade pip
pip install pypiwin32

copy /Y "%IMPORTER_PRIVATE_DIR%\DataLoader\key.txt" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\DataLoader\key.txt"

echo ***************
IF "%IMPORT_ENVIRONMENT%" == "Sandbox" (
    echo *****Sandbox Data Import Automation
    python "%IMPORTER_DIRECTORY%\..\importer_sandbox.py" %IMPORT_ENVIRONMENT% %CLIENT_TYPE% %IMPORT_MODE% %EMAIL_LIST% %IMPORT_WAITTIME% %IMPORT_NOREFRESH% %IMPORT_NOUPDATE% %IMPORT_NOEXPORTODBC% %IMPORT_NOEXPORTSF% %IMPORT_INSERTATTEMPTS% %IMPORT_ENABLE_ATTACHMENTS% %IMPORT_INTERACTIVEMODE%
) else (
    echo *****Production Data Import Automation
    python "%IMPORTER_DIRECTORY%\..\importer.py" %IMPORT_ENVIRONMENT% %CLIENT_TYPE% %IMPORT_MODE% %EMAIL_LIST% %IMPORT_WAITTIME% %IMPORT_NOREFRESH% %IMPORT_NOUPDATE% %IMPORT_NOEXPORTODBC% %IMPORT_NOEXPORTSF% %IMPORT_INSERTATTEMPTS% %IMPORT_ENABLE_ATTACHMENTS% %IMPORT_INTERACTIVEMODE%
)
echo ***************

cd %IMPORTER_PRIVATE_DIR%

:scriptexit

ENDLOCAL
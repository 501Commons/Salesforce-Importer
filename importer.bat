@echo off

:checksystem

set SF_DATALOADER=C:\Program Files (x86)\salesforce.com\Data Loader
IF NOT EXIST "%SF_DATALOADER%" (
    ECHO Error: "%SF_DATALOADER%" does not exist
    cscript importer.vbs "Salesforce Data Loader Required - a browser will open with the install info. After install press any key to continue."
    call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
    pause
    goto checksystem
)

REM Java should get installed as part of Data Loader
IF NOT EXIST "%JAVA_HOME%" (
    ECHO Error: "%JAVA_HOME%" does not exist
    cscript importer.vbs "Java not found which is installed via Salesforce Data Loader - Reinstall Data Loader and follow the instructions for Java setup.  A browser will open with the install info. After install press any key to continue."
    call explorer "https://help.salesforce.com/articleView?id=000239784&type=1"
    pause
    goto checksystem
)

IF NOT EXIST "%PYTHON_HOME%" (
    ECHO Error Python 2.7.14 Required: "%PYTHON_HOME%" does not exist
    cscript importer.vbs "Python 2.7.14 Required - a browser will open with the install info. After install press any key to continue."
    call explorer "https://www.python.org/downloads/"
    pause
    goto checksystem
)

REM IF NOT EXIST "IMPORT_DIRECTORY" (
REM    cscript importer.vbs "Error Import Directory does not exist: %IMPORT_DIRECTORY%"
REM     goto scriptexit
REM )

set PATH=%PATH%;%JAVA_HOME%;%PYTHON_HOME%;%PYTHON_HOME%\Scripts

cd "%PYTHON_HOME%"\Scripts
pip install pypiwin32

REM SANDBOX
REM xcopy "%IMPORT_DIRECTORY%" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" /s /y /i
REM copy /Y "%IMPORTER_PRIVATE_DIR%\DataLoader\key.txt" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\DataLoader\key.txt"
REM python "%IMPORTER_DIRECTORY%\..\importer.py" Sandbox %CLIENT_TYPE% %1 %EMAIL_LIST% -waittime 10 -insertattempts 1 -noupdate -noexportodbc -noexportsf

REM PRODUCTION
xcopy "%IMPORT_DIRECTORY%" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\Incoming" /s /y /i
copy /Y "%IMPORTER_PRIVATE_DIR%\DataLoader\key.txt" "%IMPORTER_DIRECTORY%\%CLIENT_TYPE%\DataLoader\key.txt"
python "%IMPORTER_DIRECTORY%\..\importer.py" Prod %CLIENT_TYPE% Import %EMAIL_LIST% -waittime 10 -insertattempts 1 -noupdate -noexportodbc -noexportsf

cd %IMPORTER_PRIVATE_DIR%

:scriptexit
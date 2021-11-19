@echo off

set EXE_PATH=%~dp0
set DATALOADER_VERSION=53.0.0

IF "%JAVA_HOME%" == "" (
    echo To run encrypt.bat, set the JAVA_HOME environment variable to the directory where the Java Runtime Environment ^(JRE^) is installed.
) ELSE (
    IF NOT EXIST "%JAVA_HOME%" (
        echo We couldn't find the Java Runtime Environment ^(JRE^) in directory "%JAVA_HOME%". To run process.bat, set the JAVA_HOME environment variable to the directory where the JRE is installed.
    ) ELSE (
        "%JAVA_HOME%\bin\java"  -cp "%EXE_PATH%\..\dataloader-%DATALOADER_VERSION%-uber.jar" com.salesforce.dataloader.security.EncryptionUtil %*

    )
)






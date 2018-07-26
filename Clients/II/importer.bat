@echo off

REM IMPORT_ENVIRONMENT can be Prod or Sandbox
set IMPORT_ENVIRONMENT=Prod

set IMPORT_WAITTIME=-waittime 300
set IMPORT_NOUPDATE=
set IMPORT_NOEXPORTODBC=-noexportodbc
set IMPORT_NOEXPORTSF=-noexportsf
set IMPORT_INSERTATTEMPTS=
set IMPORT_INTERACTIVEMODE=-interactivemode

call "%IMPORTER_DIRECTORY%\..\importer.bat" 
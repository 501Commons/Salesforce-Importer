@echo off

set IMPORT_WAITTIME=-waittime 300
set IMPORT_NOUPDATE=
set IMPORT_NOEXPORTODBC=-noexportodbc
set IMPORT_NOEXPORTSF=-noexportsf
set IMPORT_INSERTATTEMPTS=
set IMPORT_INTERACTIVEMODE=-interactivemode

call "%IMPORTER_DIRECTORY%\..\importer.bat" 
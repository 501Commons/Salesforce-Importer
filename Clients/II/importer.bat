@echo off

set IMPORT_WAITTIME=-waittime 30
set IMPORT_NOREFRESH=
set IMPORT_NOUPDATE=
set IMPORT_NOEXPORTODBC=-noexportodbc
set IMPORT_NOEXPORTSF=
set IMPORT_INSERTATTEMPTS=1
set IMPORT_ENABLE_ATTACHMENTS=
set IMPORT_INTERACTIVEMODE=
set IMPORT_DISPLAYALERTS=

call "%IMPORTER_DIRECTORY%\..\importer.bat" 
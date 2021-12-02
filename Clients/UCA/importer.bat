@echo off

set IMPORT_WAITTIME=-waittime 300
set IMPORT_NOREFRESH=
set IMPORT_NOUPDATE=-noupdate
set IMPORT_NOEXPORTODBC=-noexportodbc
set IMPORT_NOEXPORTSF=-noexportsf
set IMPORT_INSERTATTEMPTS=
set IMPORT_EMAILATTACHMENTS=-emailattachments
set IMPORT_INTERACTIVEMODE=-interactivemode
set IMPORT_DISPLAYALERTS=-displayalerts

call "%IMPORTER_DIRECTORY%\..\importer.bat" 
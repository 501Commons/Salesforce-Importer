@echo off

set IMPORT_WAITTIME=-waittime 30
set IMPORT_NOUPDATE=-noupdate
set IMPORT_NOEXPORTODBC=-noexportodbc
set IMPORT_NOEXPORTSF=-noexportsf
set IMPORT_INSERTATTEMPTS=-insertattempts 1

call "%IMPORTER_DIRECTORY%\..\importer.bat" 
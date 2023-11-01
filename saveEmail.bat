echo Starte Sicherung der Emails
set outlookfile=C:\Users\%username%\Documents\Outlook-Dateien\mail@t-online.de.pst
set backuppath=C:\Users\%username%\OneDrive\Backup
IF NOT EXIST %backuppath%\lastSave.txt goto startSave
FOR /F "tokens=* delims=" %%LastSave in (%backuppath%\lastSave.txt) DO echo %%LastSave
IF %%LastSave == %date% goto end

:startSave
echo %date% %time% Starte Sicherung der Emails > %backuppath%\backup.txt
copy %outlookfile% %backuppath%\Emails\%date%mail@t-online.de.pst
echo Kopieren der Emails abgeschlossen!
echo %date% %time% Sicherung abgeschlossen > %backuppath%\backup.txt

echo Löschen alte Sicherungen
forfiles /p %backuppath%\Emails /s /m *.pst /d -7 /c "cmd /c del @path"
echo Löschen alte Sicherungen abgeschlossen!
echo %date% %time% Löschen alte Sicherungen abgeschlossen >> %backuppath%\backup.txt
echo Sicherung der Emails abgeschlossen!
IF EXIST %backuppath%\lastSave.txt del %backuppath%\lastSave.txt
echo %date% >> %backuppath%\lastSave.txt

:end

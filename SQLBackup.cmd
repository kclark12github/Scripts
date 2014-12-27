sqlcmd -S .\SQLEXPRESS -E -Q "EXEC sp_BackupDatabases @backupLocation='\\Alpha\Backups\Databases\SQLEXPRESS2014\', @backupType='F'"

-- E-05 | Politica de respaldo: historial de backups (completo, diferencial, log)
SELECT TOP 3 database_name BaseDatos,
       CASE type WHEN 'D' THEN 'COMPLETO' WHEN 'I' THEN 'DIFERENCIAL' WHEN 'L' THEN 'LOG' END Tipo,
       backup_finish_date Fin, CAST(backup_size/1048576.0 AS DECIMAL(10,2)) TamanoMB, CAST(compressed_backup_size/1048576.0 AS DECIMAL(10,2)) ComprimidoMB
FROM msdb.dbo.backupset WHERE database_name='DistribuidoraAndina' ORDER BY backup_finish_date DESC;

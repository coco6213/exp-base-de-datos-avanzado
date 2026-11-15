/* =====================================================================
   SCRIPT 09: Politica de respaldo (backup) y prueba de recuperacion.
   EN SIMPLE: un backup es una "foto" de la base guardada aparte.
     - COMPLETO     : foto de toda la base          -> domingo 01:00
     - DIFERENCIAL  : solo lo que cambio desde el completo -> diario 01:00
     - LOG          : cada movimiento (transaccion) -> cada 15 minutos
   Asi, si algo falla, se pierde como maximo 15 minutos de informacion.
   ===================================================================== */
USE master;
GO
DECLARE @ruta NVARCHAR(400) = CAST(SERVERPROPERTY('InstanceDefaultBackupPath') AS NVARCHAR(400)) + N'\';
DECLARE @f NVARCHAR(400);

SET @f = @ruta + N'DistribuidoraAndina_FULL.bak';
BACKUP DATABASE DistribuidoraAndina TO DISK = @f WITH INIT, COMPRESSION, CHECKSUM, NAME = N'Backup COMPLETO';

SET @f = @ruta + N'DistribuidoraAndina_DIFF.bak';
BACKUP DATABASE DistribuidoraAndina TO DISK = @f WITH DIFFERENTIAL, INIT, COMPRESSION, CHECKSUM, NAME = N'Backup DIFERENCIAL';

SET @f = @ruta + N'DistribuidoraAndina_LOG.trn';
BACKUP LOG DistribuidoraAndina TO DISK = @f WITH INIT, COMPRESSION, CHECKSUM, NAME = N'Backup de LOG';

-- Verificacion: el archivo se puede leer y no esta danado (no restaura, solo comprueba)
SET @f = @ruta + N'DistribuidoraAndina_FULL.bak';
RESTORE VERIFYONLY FROM DISK = @f WITH CHECKSUM;
GO
-- Evidencia: historial de respaldos registrado por SQL Server
SELECT TOP 3 bs.database_name AS BaseDatos,
       CASE bs.type WHEN 'D' THEN 'COMPLETO' WHEN 'I' THEN 'DIFERENCIAL' WHEN 'L' THEN 'LOG' END AS Tipo,
       bs.backup_start_date AS Inicio, bs.backup_finish_date AS Fin,
       CAST(bs.backup_size / 1048576.0 AS DECIMAL(10,2)) AS TamanoMB,
       CAST(bs.compressed_backup_size / 1048576.0 AS DECIMAL(10,2)) AS ComprimidoMB,
       bmf.physical_device_name AS Archivo
FROM msdb.dbo.backupset bs JOIN msdb.dbo.backupmediafamily bmf ON bmf.media_set_id = bs.media_set_id
WHERE bs.database_name = 'DistribuidoraAndina'
ORDER BY bs.backup_finish_date DESC;
GO

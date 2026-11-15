-- E-06b | Trigger de auditoria: quien cambio que y cuando
USE DistribuidoraAndina;
UPDATE dbo.Producto SET PrecioUnitario = PrecioUnitario + 0.10 WHERE ProductoID = 2;
SELECT TOP 4 AuditID, Tabla, Operacion, Usuario, Fecha,
       JSON_VALUE(DatosAnteriores,'$[0].PrecioUnitario') PrecioAntes, JSON_VALUE(DatosNuevos,'$[0].PrecioUnitario') PrecioDespues
FROM dbo.AuditoriaLog ORDER BY AuditID DESC;

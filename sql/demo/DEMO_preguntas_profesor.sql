/* =====================================================================
   DEMOSTRACIÓN EN VIVO – DataShield 360 (Distribuidora Andina)
   CÓMO USAR: seleccione (sombree) SOLO el bloque de la pregunta y presione F5.
   Cada bloque es independiente. Los bloques 1 a 6 solo LEEN datos.
   ===================================================================== */

/* ---------- 1. ¿Cuántos datos tienen? ---------- */
SELECT 'Ventas' AS Tabla, COUNT(*) AS Filas FROM DistribuidoraAndina.dbo.Venta
UNION ALL SELECT 'Detalle de ventas', COUNT(*) FROM DistribuidoraAndina.dbo.DetalleVenta
UNION ALL SELECT 'Clientes', COUNT(*) FROM DistribuidoraAndina.dbo.Cliente
UNION ALL SELECT 'Productos', COUNT(*) FROM DistribuidoraAndina.dbo.Producto
UNION ALL SELECT 'Data Warehouse (FactVentas)', COUNT(*) FROM DW_Andina.dbo.FactVentas;

/* ---------- 2. ¿Qué tablas tiene la base y cómo se relacionan? ---------- */
SELECT t.name AS Tabla, fk.name AS LlaveForanea, OBJECT_NAME(fk.referenced_object_id, DB_ID('DistribuidoraAndina')) AS ApuntaA
FROM DistribuidoraAndina.sys.tables t
LEFT JOIN DistribuidoraAndina.sys.foreign_keys fk ON fk.parent_object_id = t.object_id
ORDER BY t.name;

/* ---------- 3. ¿El ETL cuadra con la base original? (reconciliación) ---------- */
SELECT (SELECT SUM(dv.Subtotal) FROM DistribuidoraAndina.dbo.DetalleVenta dv
        JOIN DistribuidoraAndina.dbo.Venta v ON v.VentaID = dv.VentaID
        WHERE v.Estado = 'EMITIDA' AND v.VentaID <= 200001) AS ImporteOLTP,
       (SELECT SUM(Importe) FROM DW_Andina.dbo.FactVentas) AS ImporteDW;
SELECT * FROM DW_Andina.dbo.ETL_Log;                       -- extraídas / rechazadas / cargadas

/* ---------- 4. Los KPI del dashboard (2025 vs meta) ---------- */
WITH a AS (SELECT Anio, SUM(Importe) Ventas, SUM(Unidades) Unidades FROM DW_Andina.dbo.vw_Dashboard GROUP BY Anio)
SELECT 'KPI 1 Crecimiento 2025' AS KPI, CAST(100.0 * (a25.Ventas / a24.Ventas - 1) AS DECIMAL(6,2)) AS Real, '>= +5 %' AS Meta,
       CASE WHEN a25.Ventas / a24.Ventas - 1 >= 0.05 THEN 'CUMPLE' ELSE 'NO CUMPLE' END AS Estado
FROM a a25 JOIN a a24 ON a25.Anio = 2025 AND a24.Anio = 2024
UNION ALL
SELECT 'KPI 2 Cumplimiento meta ventas', CAST(100.0 * a25.Ventas / (a24.Ventas * 1.05) AS DECIMAL(6,2)), '100 %',
       CASE WHEN a25.Ventas >= a24.Ventas * 1.05 THEN 'CUMPLE' ELSE 'NO CUMPLE' END
FROM a a25 JOIN a a24 ON a25.Anio = 2025 AND a24.Anio = 2024
UNION ALL
SELECT 'KPI 3 Precio promedio 2025', CAST(a25.Ventas / a25.Unidades AS DECIMAL(6,2)), '>= S/ 24.00',
       CASE WHEN a25.Ventas / a25.Unidades >= 24 THEN 'CUMPLE' ELSE 'NO CUMPLE' END
FROM a a25 WHERE a25.Anio = 2025;

/* ---------- 5. Consultas OLAP en el Data Warehouse ---------- */
-- 5a. Ventas por región y año (slice and dice)
SELECT s.Region, t.Anio, SUM(f.Importe) AS Ventas
FROM DW_Andina.dbo.FactVentas f
JOIN DW_Andina.dbo.DimSede s ON s.SedeKey = f.SedeKey
JOIN DW_Andina.dbo.DimTiempo t ON t.TiempoKey = f.TiempoKey
WHERE t.Anio BETWEEN 2023 AND 2025
GROUP BY s.Region, t.Anio ORDER BY s.Region, t.Anio;
-- 5b. Top 5 productos
SELECT TOP 5 p.Nombre, SUM(f.Importe) AS Ventas
FROM DW_Andina.dbo.FactVentas f JOIN DW_Andina.dbo.DimProducto p ON p.ProductoKey = f.ProductoKey
GROUP BY p.Nombre ORDER BY Ventas DESC;
-- 5c. Campaña Nov-Dic vs resto del año (venta promedio por mes)
SELECT CASE WHEN t.EsCampana = 1 THEN 'Campaña Nov-Dic' ELSE 'Resto del año' END AS Periodo,
       SUM(f.Importe) / COUNT(DISTINCT t.Anio * 100 + t.Mes) AS VentaPromedioMes
FROM DW_Andina.dbo.FactVentas f JOIN DW_Andina.dbo.DimTiempo t ON t.TiempoKey = f.TiempoKey
WHERE t.Anio BETWEEN 2023 AND 2025
GROUP BY t.EsCampana;

/* ---------- 6. Seguridad: el analista NO puede ver datos personales ---------- */
USE DistribuidoraAndina;
EXECUTE AS USER = 'usr_analista';
SELECT TOP 3 ClienteID, RazonSocial, Segmento FROM dbo.Cliente;   -- permitido
BEGIN TRY SELECT TOP 3 Email FROM dbo.Cliente; END TRY            -- debe ser denegado
BEGIN CATCH SELECT 'ACCESO DENEGADO' AS Resultado, ERROR_MESSAGE() AS Mensaje; END CATCH
REVERT;

/* ---------- 7. Procedimiento y transacción (ESCRIBE: agrega 1 venta) ---------- */
USE DistribuidoraAndina;
EXEC dbo.usp_RegistrarVenta @ClienteID = 30, @VendedorID = 1, @SedeID = 1, @Detalle = N'[{"ProductoID":9,"Cantidad":1}]';      -- se guarda (COMMIT)
EXEC dbo.usp_RegistrarVenta @ClienteID = 30, @VendedorID = 1, @SedeID = 1, @Detalle = N'[{"ProductoID":9,"Cantidad":999999}]'; -- sin stock: ROLLBACK
SELECT TOP 3 VentaID, Fecha, ClienteID, Total, Estado FROM dbo.Venta ORDER BY VentaID DESC;

/* ---------- 8. Trigger de auditoría (ESCRIBE: cambia un precio en 0.01) ---------- */
USE DistribuidoraAndina;
UPDATE dbo.Producto SET PrecioUnitario = PrecioUnitario + 0.01 WHERE ProductoID = 3;
SELECT TOP 3 AuditID, Tabla, Operacion, Usuario, Fecha,
       JSON_VALUE(DatosAnteriores, '$[0].PrecioUnitario') AS PrecioAntes,
       JSON_VALUE(DatosNuevos, '$[0].PrecioUnitario') AS PrecioDespues
FROM dbo.AuditoriaLog ORDER BY AuditID DESC;

/* ---------- 9. Respaldos (backup) realizados ---------- */
SELECT TOP 3 database_name AS BaseDatos,
       CASE type WHEN 'D' THEN 'COMPLETO' WHEN 'I' THEN 'DIFERENCIAL' WHEN 'L' THEN 'LOG' END AS Tipo,
       backup_finish_date AS Fecha, CAST(backup_size / 1048576.0 AS DECIMAL(10,2)) AS TamanoMB
FROM msdb.dbo.backupset WHERE database_name = 'DistribuidoraAndina' ORDER BY backup_finish_date DESC;

/* ---------- 10. Prueba de viabilidad: ¿por qué el Data Warehouse es más rápido? ---------- */
SELECT Escenario, ConsultaID, CAST(AVG(TiempoMs) AS DECIMAL(10,1)) AS TiempoPromedio_ms, AVG(Lecturas) AS PaginasLeidas
FROM LabExperimento.dbo.Resultado WHERE TipoCorrida = 'CALIENTE'
GROUP BY Escenario, ConsultaID ORDER BY ConsultaID, Escenario;

/* ---------- 11. Ver la definición de un procedimiento o trigger ---------- */
EXEC DistribuidoraAndina.sys.sp_helptext 'dbo.usp_RegistrarVenta';
-- EXEC DistribuidoraAndina.sys.sp_helptext 'dbo.trg_Auditoria_Producto';

/* =====================================================================
   SCRIPT 10: Vista que alimenta el dashboard BI (Excel).
   EN SIMPLE: junta la tabla de hechos con sus dimensiones y resume por
   mes, producto, sede y segmento. El dashboard se "conecta" a esta vista
   y con un clic en "Actualizar" trae los datos nuevos del Data Warehouse.
   ===================================================================== */
USE DW_Andina;
GO
CREATE OR ALTER VIEW dbo.vw_Dashboard AS
SELECT t.Anio,
       t.Mes,
       CONCAT(t.Anio, '-', RIGHT('0' + CAST(t.Mes AS VARCHAR(2)), 2)) AS Periodo,
       t.Trimestre,
       CASE WHEN t.EsCampana = 1 THEN 'CAMPANA NOV-DIC' ELSE 'REGULAR' END AS Temporada,
       s.Region, s.Sede,
       p.Categoria, p.Nombre AS Producto,
       c.Segmento,
       SUM(f.Cantidad)            AS Unidades,
       SUM(f.Importe)             AS Importe,
       COUNT(DISTINCT f.VentaID)  AS Ventas
FROM dbo.FactVentas f
JOIN dbo.DimTiempo   t ON t.TiempoKey   = f.TiempoKey
JOIN dbo.DimSede     s ON s.SedeKey     = f.SedeKey
JOIN dbo.DimProducto p ON p.ProductoKey = f.ProductoKey
JOIN dbo.DimCliente  c ON c.ClienteKey  = f.ClienteKey
WHERE t.Anio BETWEEN 2023 AND 2025            -- periodo de analisis (excluye la venta de prueba 2026)
GROUP BY t.Anio, t.Mes, t.Trimestre, t.EsCampana, s.Region, s.Sede, p.Categoria, p.Nombre, c.Segmento;
GO
-- Rol de solo lectura para la herramienta BI
IF DATABASE_PRINCIPAL_ID('rol_lector_bi') IS NULL CREATE ROLE rol_lector_bi;
GRANT SELECT ON dbo.vw_Dashboard TO rol_lector_bi;
GO
SELECT COUNT(*) AS FilasVista, SUM(Importe) AS ImporteTotal, SUM(Unidades) AS Unidades FROM dbo.vw_Dashboard;
-- KPIs por anio (para validar el dashboard)
SELECT Anio, SUM(Importe) Importe, SUM(Unidades) Unidades,
       CAST(100.0 * (SUM(Importe) - LAG(SUM(Importe)) OVER (ORDER BY Anio)) / LAG(SUM(Importe)) OVER (ORDER BY Anio) AS DECIMAL(6,2)) CrecimientoPct
FROM dbo.vw_Dashboard GROUP BY Anio ORDER BY Anio;
SELECT Region, SUM(Importe) Importe FROM dbo.vw_Dashboard GROUP BY Region ORDER BY Importe DESC;
SELECT TOP 5 Producto, SUM(Importe) Importe FROM dbo.vw_Dashboard GROUP BY Producto ORDER BY Importe DESC;
SELECT Temporada, SUM(Importe) / COUNT(DISTINCT Periodo) ImportePromedioMes FROM dbo.vw_Dashboard GROUP BY Temporada;
GO

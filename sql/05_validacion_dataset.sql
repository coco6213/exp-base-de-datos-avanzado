/* =====================================================================
   SCRIPT 05: Validacion del dataset (antes de experimentar).
   EN SIMPLE: revisamos que los datos esten "sanos": que no falten,
   que no haya repetidos, que las relaciones esten completas y que los
   totales cuadren. Si algo falla aqui, el experimento no seria confiable.
   ===================================================================== */
USE DistribuidoraAndina;
GO
SET NOCOUNT ON;

-- V1. Volumen del dataset
SELECT 'V1 Volumen' AS Prueba, 'Venta' Tabla, COUNT(*) Filas FROM dbo.Venta
UNION ALL SELECT 'V1 Volumen','DetalleVenta', COUNT(*) FROM dbo.DetalleVenta
UNION ALL SELECT 'V1 Volumen','Cliente', COUNT(*) FROM dbo.Cliente
UNION ALL SELECT 'V1 Volumen','Producto', COUNT(*) FROM dbo.Producto;

-- V2..V8: cada prueba devuelve el numero de problemas encontrados y si pasa o no
;WITH chequeos AS (
  SELECT 'V2 Ventas sin detalle'              Prueba, COUNT(*) Problemas, 'Debe ser 0' Criterio
  FROM dbo.Venta v WHERE NOT EXISTS (SELECT 1 FROM dbo.DetalleVenta d WHERE d.VentaID = v.VentaID)
  UNION ALL
  SELECT 'V3 Detalle huerfano (sin venta)', COUNT(*), 'Debe ser 0'
  FROM dbo.DetalleVenta d WHERE NOT EXISTS (SELECT 1 FROM dbo.Venta v WHERE v.VentaID = d.VentaID)
  UNION ALL
  SELECT 'V4 RUC duplicados', COUNT(*) - COUNT(DISTINCT RUC), 'Debe ser 0' FROM dbo.Cliente
  UNION ALL
  SELECT 'V5 Total cabecera <> suma detalle', COUNT(*), 'Debe ser 0'
  FROM dbo.Venta v JOIN (SELECT VentaID, SUM(Subtotal) s FROM dbo.DetalleVenta GROUP BY VentaID) d
       ON d.VentaID = v.VentaID WHERE v.Total <> d.s
  UNION ALL
  SELECT 'V6 Fechas fuera de rango 2023-2025', COUNT(*), 'Debe ser 0'
  FROM dbo.Venta WHERE Fecha < '2023-01-01' OR Fecha >= '2026-01-01'
  UNION ALL
  SELECT 'V7 Cantidades o precios invalidos', COUNT(*), 'Debe ser 0'
  FROM dbo.DetalleVenta WHERE Cantidad <= 0 OR PrecioUnitario <= 0
  UNION ALL
  SELECT 'V8 Cliente de otra region que la sede', COUNT(*), 'Debe ser 0'
  FROM dbo.Venta v JOIN dbo.Cliente c ON c.ClienteID = v.ClienteID JOIN dbo.Sede s ON s.SedeID = v.SedeID
  WHERE c.RegionID <> s.RegionID
  UNION ALL
  -- V13 (agregada tras detectar sesgo en v1): toda combinacion region x segmento debe tener clientes
  SELECT 'V13 Combinaciones region x segmento vacias', 9 - COUNT(DISTINCT CAST(RegionID AS VARCHAR) + Segmento), 'Debe ser 0'
  FROM dbo.Cliente
)
SELECT Prueba, Problemas, Criterio, CASE WHEN Problemas = 0 THEN 'PASA' ELSE 'FALLA' END Resultado FROM chequeos;

-- V9..V11: datos "sucios" esperados (los debe corregir el ETL, no son errores del generador)
SELECT 'V9 Clientes sin email'               Hallazgo, COUNT(*) Cantidad, 'Regla ETL R4' Tratamiento FROM dbo.Cliente WHERE Email IS NULL
UNION ALL
SELECT 'V10 Productos con nombre sucio', COUNT(*), 'Regla ETL R1' FROM dbo.Producto WHERE Nombre <> UPPER(LTRIM(RTRIM(Nombre))) COLLATE Latin1_General_CS_AS
UNION ALL
SELECT 'V11 Ventas anuladas', COUNT(*), 'Regla ETL R2' FROM dbo.Venta WHERE Estado = 'ANULADA';

-- V0: "huella digital" del dataset. Si se regenera con el script 02 y la huella es igual,
--     el dataset es identico (prueba de reproducibilidad).
SELECT 'V0 Huella dataset' AS Prueba,
       (SELECT CHECKSUM_AGG(CHECKSUM(VentaID, Fecha, ClienteID, VendedorID, SedeID, Total, Estado)) FROM dbo.Venta WHERE VentaID <= 200000) AS HuellaVenta,
       (SELECT CHECKSUM_AGG(CHECKSUM(DetalleID, VentaID, ProductoID, Cantidad, PrecioUnitario)) FROM dbo.DetalleVenta WHERE VentaID <= 200000) AS HuellaDetalle,
       (SELECT CHECKSUM_AGG(CHECKSUM(ClienteID, RUC, Segmento, RegionID)) FROM dbo.Cliente) AS HuellaCliente;

-- V12: distribucion por region y estacionalidad (perfil del dataset)
SELECT r.Nombre Region, COUNT(*) Ventas, CAST(100.0*COUNT(*)/SUM(COUNT(*)) OVER() AS DECIMAL(5,2)) Porcentaje
FROM dbo.Venta v JOIN dbo.Sede s ON s.SedeID = v.SedeID JOIN dbo.Region r ON r.RegionID = s.RegionID
GROUP BY r.Nombre ORDER BY Ventas DESC;

SELECT CASE WHEN MONTH(Fecha) IN (11,12) THEN 'Nov-Dic (campana)' ELSE 'Resto del anio' END Periodo,
       COUNT(*) Ventas,
       CAST(COUNT(*) / (CASE WHEN MAX(MONTH(Fecha)) <= 12 AND MIN(MONTH(Fecha)) >= 11 THEN 6.0 ELSE 30.0 END) AS INT) VentasPromedioPorMes  -- 6 meses de campana vs 30 meses normales en 3 anios
FROM dbo.Venta GROUP BY CASE WHEN MONTH(Fecha) IN (11,12) THEN 'Nov-Dic (campana)' ELSE 'Resto del anio' END;
GO

/* =====================================================================
   SCRIPT 07: EXPERIMENTO - Comparacion de 3 escenarios de consulta.
   PREGUNTA : Cuanto mejora el tiempo de respuesta de los reportes de ventas
              si pasamos de la base diaria "tal cual" (E1) a la misma base con
              indices (E2) y a un Data Warehouse con indice columnar (E3)?
   ESCENARIOS:
     E1 = OLTP base   : solo llaves primarias (situacion inicial de la empresa)
     E2 = OLTP + IX   : se agregan 4 indices no agrupados ("indices de libro")
     E3 = DW estrella : tabla de hechos con indice COLUMNSTORE
   CONSULTAS (las 4 preguntas que hace la gerencia):
     Q1 Ventas por region y mes (2025)        -> analisis OLAP tiempo x region
     Q2 Top 10 productos mas vendidos          -> ranking
     Q3 Ventas por categoria y anio            -> OLAP producto x tiempo
     Q4 Compras de un cliente en 2024          -> consulta puntual/operativa
   METRICAS : tiempo (ms), CPU (ms), lecturas logicas (paginas de 8 KB leidas).
   PROCEDIMIENTO: por cada consulta -> 1 corrida en FRIO (memoria vaciada) +
              10 corridas en CALIENTE. Se guarda todo en LabExperimento.dbo.Resultado.
   VALIDEZ  : se compara un CHECKSUM del resultado: los 3 escenarios deben
              devolver EXACTAMENTE los mismos numeros.
   ===================================================================== */
USE master;
GO
IF DB_ID('LabExperimento') IS NULL CREATE DATABASE LabExperimento;
GO
USE LabExperimento;
GO
DROP TABLE IF EXISTS dbo.Resultado, dbo.Consulta, dbo.Almacenamiento;
CREATE TABLE dbo.Consulta (
    Escenario   CHAR(2)      NOT NULL,
    ConsultaID  CHAR(2)      NOT NULL,
    Descripcion VARCHAR(100) NOT NULL,
    SqlTexto    NVARCHAR(MAX) NOT NULL,
    CONSTRAINT PK_Consulta PRIMARY KEY (Escenario, ConsultaID)
);
CREATE TABLE dbo.Resultado (
    ID          INT IDENTITY(1,1) PRIMARY KEY,
    Escenario   CHAR(2)  NOT NULL,
    ConsultaID  CHAR(2)  NOT NULL,
    Corrida     INT      NOT NULL,          -- 0 = frio, 1..10 = caliente
    TipoCorrida VARCHAR(8) NOT NULL,
    TiempoMs    DECIMAL(10,3) NOT NULL,
    CpuMs       INT      NOT NULL,
    Lecturas    BIGINT   NOT NULL,
    Filas       INT      NOT NULL,
    ChecksumRes INT      NULL,
    FechaHora   DATETIME2(0) NOT NULL DEFAULT SYSDATETIME()
);
CREATE TABLE dbo.Almacenamiento (Escenario CHAR(2), Objeto VARCHAR(60), Filas BIGINT, EspacioMB DECIMAL(10,2));
GO

/* ---------- Catalogo de consultas (mismo resultado logico en los 3 escenarios) ---------- */
DECLARE @oltp TABLE (ConsultaID CHAR(2), Descripcion VARCHAR(100), SqlTexto NVARCHAR(MAX));
INSERT @oltp VALUES
('Q1','Ventas por region y mes (2025)',
 N'SELECT CAST(r.Nombre AS VARCHAR(50)) Region, CAST(MONTH(v.Fecha) AS INT) Mes, SUM(d.Subtotal) Importe
   FROM DistribuidoraAndina.dbo.Venta v
   JOIN DistribuidoraAndina.dbo.DetalleVenta d ON d.VentaID = v.VentaID
   JOIN DistribuidoraAndina.dbo.Sede s ON s.SedeID = v.SedeID
   JOIN DistribuidoraAndina.dbo.Region r ON r.RegionID = s.RegionID
   WHERE v.Estado = ''EMITIDA'' AND v.Fecha >= ''2025-01-01'' AND v.Fecha < ''2026-01-01''
   GROUP BY r.Nombre, MONTH(v.Fecha)'),
('Q2','Top 10 productos por importe',
 N'SELECT TOP (10) d.ProductoID, SUM(d.Subtotal) Importe
   FROM DistribuidoraAndina.dbo.DetalleVenta d
   JOIN DistribuidoraAndina.dbo.Venta v ON v.VentaID = d.VentaID
   WHERE v.Estado = ''EMITIDA''
   GROUP BY d.ProductoID ORDER BY Importe DESC'),
('Q3','Ventas por categoria y anio',
 N'SELECT CAST(c.Nombre AS VARCHAR(60)) Categoria, CAST(YEAR(v.Fecha) AS INT) Anio, SUM(d.Subtotal) Importe
   FROM DistribuidoraAndina.dbo.DetalleVenta d
   JOIN DistribuidoraAndina.dbo.Venta v ON v.VentaID = d.VentaID
   JOIN DistribuidoraAndina.dbo.Producto p ON p.ProductoID = d.ProductoID
   JOIN DistribuidoraAndina.dbo.Categoria c ON c.CategoriaID = p.CategoriaID
   WHERE v.Estado = ''EMITIDA''
   GROUP BY c.Nombre, YEAR(v.Fecha)'),
('Q4','Compras del cliente 1234 en 2024',
 N'SELECT v.VentaID, SUM(d.Subtotal) Importe
   FROM DistribuidoraAndina.dbo.Venta v
   JOIN DistribuidoraAndina.dbo.DetalleVenta d ON d.VentaID = v.VentaID
   WHERE v.ClienteID = 1234 AND v.Estado = ''EMITIDA''
     AND v.Fecha >= ''2024-01-01'' AND v.Fecha < ''2025-01-01''
   GROUP BY v.VentaID');

INSERT dbo.Consulta SELECT 'E1', * FROM @oltp;   -- E1 y E2 usan el mismo SQL;
INSERT dbo.Consulta SELECT 'E2', * FROM @oltp;   -- lo unico que cambia son los indices

INSERT dbo.Consulta VALUES
('E3','Q1','Ventas por region y mes (2025)',
 N'SELECT CAST(s.Region AS VARCHAR(50)) Region, CAST(t.Mes AS INT) Mes, SUM(f.Importe) Importe
   FROM DW_Andina.dbo.FactVentas f
   JOIN DW_Andina.dbo.DimSede s ON s.SedeKey = f.SedeKey
   JOIN DW_Andina.dbo.DimTiempo t ON t.TiempoKey = f.TiempoKey
   WHERE t.Anio = 2025
   GROUP BY s.Region, t.Mes'),
('E3','Q2','Top 10 productos por importe',
 N'SELECT TOP (10) p.ProductoID, SUM(f.Importe) Importe
   FROM DW_Andina.dbo.FactVentas f
   JOIN DW_Andina.dbo.DimProducto p ON p.ProductoKey = f.ProductoKey
   GROUP BY p.ProductoID ORDER BY Importe DESC'),
('E3','Q3','Ventas por categoria y anio',
 N'SELECT CAST(p.Categoria AS VARCHAR(60)) Categoria, CAST(t.Anio AS INT) Anio, SUM(f.Importe) Importe
   FROM DW_Andina.dbo.FactVentas f
   JOIN DW_Andina.dbo.DimProducto p ON p.ProductoKey = f.ProductoKey
   JOIN DW_Andina.dbo.DimTiempo t ON t.TiempoKey = f.TiempoKey
   GROUP BY p.Categoria, t.Anio'),
('E3','Q4','Compras del cliente 1234 en 2024',
 N'SELECT f.VentaID, SUM(f.Importe) Importe
   FROM DW_Andina.dbo.FactVentas f
   JOIN DW_Andina.dbo.DimCliente c ON c.ClienteKey = f.ClienteKey
   WHERE c.ClienteID = 1234 AND f.TiempoKey BETWEEN 20240101 AND 20241231
   GROUP BY f.VentaID');
GO

/* ---------- Procedimiento que mide (cronometro + contador de CPU y lecturas) ---------- */
CREATE OR ALTER PROCEDURE dbo.usp_Medir @Escenario CHAR(2), @Corridas INT = 10
AS
BEGIN
    SET NOCOUNT ON;
    DECLARE @cid CHAR(2), @sql NVARCHAR(MAX), @wrap NVARCHAR(MAX), @i INT,
            @c0 INT, @r0 BIGINT, @t0 DATETIME2(7), @c1 INT, @r1 BIGINT, @f INT, @chk INT;
    DECLARE cur CURSOR LOCAL FAST_FORWARD FOR
        SELECT ConsultaID, SqlTexto FROM dbo.Consulta WHERE Escenario = @Escenario ORDER BY ConsultaID;
    OPEN cur; FETCH NEXT FROM cur INTO @cid, @sql;
    WHILE @@FETCH_STATUS = 0
    BEGIN
        -- la consulta se "envuelve" para contar filas y calcular un checksum sin mostrar resultados
        SET @wrap = N'SELECT @f = COUNT(*), @chk = CHECKSUM_AGG(CHECKSUM(*)) FROM (' + @sql + N') q;';
        SET @i = 0;
        WHILE @i <= @Corridas
        BEGIN
            IF @i = 0 BEGIN CHECKPOINT; DBCC DROPCLEANBUFFERS WITH NO_INFOMSGS; END   -- corrida en FRIO
            SELECT @c0 = cpu_time, @r0 = logical_reads FROM sys.dm_exec_requests WHERE session_id = @@SPID;
            SET @t0 = SYSDATETIME();                                                -- inicia cronometro
            EXEC sp_executesql @wrap, N'@f INT OUTPUT, @chk INT OUTPUT', @f OUTPUT, @chk OUTPUT;
            SELECT @c1 = cpu_time, @r1 = logical_reads FROM sys.dm_exec_requests WHERE session_id = @@SPID;
            INSERT dbo.Resultado(Escenario, ConsultaID, Corrida, TipoCorrida, TiempoMs, CpuMs, Lecturas, Filas, ChecksumRes)
            VALUES (@Escenario, @cid, @i, CASE WHEN @i = 0 THEN 'FRIA' ELSE 'CALIENTE' END,
                    DATEDIFF(MICROSECOND, @t0, SYSDATETIME()) / 1000.0, @c1 - @c0, @r1 - @r0, @f, @chk);
            SET @i += 1;
        END
        FETCH NEXT FROM cur INTO @cid, @sql;
    END
    CLOSE cur; DEALLOCATE cur;
END
GO

/* =================== ESCENARIO E1: OLTP sin indices adicionales =================== */
USE DistribuidoraAndina;
DROP INDEX IF EXISTS IX_Venta_Fecha    ON dbo.Venta;
DROP INDEX IF EXISTS IX_Venta_Cliente  ON dbo.Venta;
DROP INDEX IF EXISTS IX_Det_Venta      ON dbo.DetalleVenta;
DROP INDEX IF EXISTS IX_Det_Producto   ON dbo.DetalleVenta;
GO
INSERT LabExperimento.dbo.Almacenamiento
SELECT 'E1', OBJECT_NAME(p.object_id), SUM(CASE WHEN p.index_id IN (0,1) THEN p.rows END), SUM(a.total_pages) * 8 / 1024.0
FROM sys.partitions p JOIN sys.allocation_units a ON a.container_id = p.partition_id
WHERE p.object_id IN (OBJECT_ID('dbo.Venta'), OBJECT_ID('dbo.DetalleVenta')) GROUP BY p.object_id;
GO
EXEC LabExperimento.dbo.usp_Medir 'E1';
GO

/* =================== ESCENARIO E2: OLTP + 4 indices no agrupados =================== */
USE DistribuidoraAndina;
CREATE INDEX IX_Venta_Fecha   ON dbo.Venta(Fecha)            INCLUDE (SedeID, Estado, ClienteID);   -- filtros por fecha
CREATE INDEX IX_Venta_Cliente ON dbo.Venta(ClienteID, Fecha) INCLUDE (Estado);                      -- busqueda por cliente
CREATE INDEX IX_Det_Venta     ON dbo.DetalleVenta(VentaID)   INCLUDE (ProductoID, Subtotal);        -- union venta-detalle
CREATE INDEX IX_Det_Producto  ON dbo.DetalleVenta(ProductoID) INCLUDE (VentaID, Subtotal);          -- agrupar por producto
GO
INSERT LabExperimento.dbo.Almacenamiento
SELECT 'E2', OBJECT_NAME(p.object_id), SUM(CASE WHEN p.index_id IN (0,1) THEN p.rows END), SUM(a.total_pages) * 8 / 1024.0
FROM sys.partitions p JOIN sys.allocation_units a ON a.container_id = p.partition_id
WHERE p.object_id IN (OBJECT_ID('dbo.Venta'), OBJECT_ID('dbo.DetalleVenta')) GROUP BY p.object_id;
GO
EXEC LabExperimento.dbo.usp_Medir 'E2';
GO

/* =================== ESCENARIO E3: Data Warehouse + COLUMNSTORE =================== */
USE DW_Andina;
INSERT LabExperimento.dbo.Almacenamiento
SELECT 'E3', OBJECT_NAME(p.object_id), SUM(CASE WHEN p.index_id IN (0,1) THEN p.rows END), SUM(a.total_pages) * 8 / 1024.0
FROM sys.partitions p JOIN sys.allocation_units a ON a.container_id = p.partition_id
WHERE p.object_id = OBJECT_ID('dbo.FactVentas') GROUP BY p.object_id;
GO
EXEC LabExperimento.dbo.usp_Medir 'E3';
GO

/* =================== RESULTADOS =================== */
USE LabExperimento;
GO
-- Tabla consolidada (corridas en caliente)
SELECT Escenario, ConsultaID,
       CAST(AVG(TiempoMs) AS DECIMAL(10,2))                 TiempoProm_ms,
       CAST(STDEV(TiempoMs) AS DECIMAL(10,2))               DesvEst_ms,
       CAST(MIN(TiempoMs) AS DECIMAL(10,2))                 TiempoMin_ms,
       CAST(MAX(TiempoMs) AS DECIMAL(10,2))                 TiempoMax_ms,
       AVG(CpuMs)                                           CpuProm_ms,
       AVG(Lecturas)                                        LecturasProm,
       MAX(Filas)                                           Filas
FROM dbo.Resultado WHERE TipoCorrida = 'CALIENTE'
GROUP BY Escenario, ConsultaID ORDER BY ConsultaID, Escenario;

-- Corrida en frio (memoria vaciada)
SELECT Escenario, ConsultaID, TiempoMs TiempoFrio_ms, CpuMs, Lecturas
FROM dbo.Resultado WHERE TipoCorrida = 'FRIA' ORDER BY ConsultaID, Escenario;

-- Mejora porcentual respecto a la situacion inicial (E1)
;WITH p AS (SELECT Escenario, ConsultaID, AVG(TiempoMs) t, AVG(Lecturas) l FROM dbo.Resultado WHERE TipoCorrida='CALIENTE' GROUP BY Escenario, ConsultaID)
SELECT b.ConsultaID, x.Escenario,
       CAST(b.t AS DECIMAL(10,2)) E1_ms, CAST(x.t AS DECIMAL(10,2)) Escenario_ms,
       CAST(100.0 * (b.t - x.t) / b.t AS DECIMAL(6,1)) MejoraTiempo_pct,
       CAST(100.0 * (b.l - x.l) / NULLIF(b.l,0) AS DECIMAL(6,1)) MenosLecturas_pct,
       CAST(b.t / NULLIF(x.t,0) AS DECIMAL(8,1)) VecesMasRapido
FROM p b JOIN p x ON x.ConsultaID = b.ConsultaID AND x.Escenario <> 'E1'
WHERE b.Escenario = 'E1' ORDER BY b.ConsultaID, x.Escenario;

-- Validez: los 3 escenarios deben dar el mismo resultado (mismo checksum y filas)
SELECT ConsultaID,
       COUNT(DISTINCT CAST(ChecksumRes AS VARCHAR(20)) + '-' + CAST(Filas AS VARCHAR(10))) ResultadosDistintos,
       CASE WHEN COUNT(DISTINCT CAST(ChecksumRes AS VARCHAR(20)) + '-' + CAST(Filas AS VARCHAR(10))) = 1
            THEN 'CONSISTENTE' ELSE 'REVISAR' END Validacion
FROM dbo.Resultado GROUP BY ConsultaID;

-- Espacio en disco por escenario
-- (correccion: en el columnstore la union con allocation_units duplica el conteo de filas; se toma el conteo real)
UPDATE dbo.Almacenamiento SET Filas = (SELECT COUNT_BIG(*) FROM DW_Andina.dbo.FactVentas) WHERE Escenario = 'E3';
SELECT * FROM dbo.Almacenamiento ORDER BY Escenario;
GO

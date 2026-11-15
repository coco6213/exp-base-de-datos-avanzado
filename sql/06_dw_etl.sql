/* =====================================================================
   SCRIPT 06: Data Warehouse (esquema estrella) + proceso ETL.
   EN SIMPLE: el Data Warehouse es una "copia ordenada para analizar".
   En el centro esta la tabla de HECHOS (cada linea vendida) y alrededor
   las DIMENSIONES (cuando, que producto, que cliente, que sede).
   ETL = Extraer (de la base diaria) -> Transformar (limpiar) -> Cargar (al DW).
   ===================================================================== */
USE master;
GO
IF DB_ID('DW_Andina') IS NOT NULL
BEGIN
    ALTER DATABASE DW_Andina SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE DW_Andina;
END
GO
CREATE DATABASE DW_Andina;
GO
ALTER DATABASE DW_Andina SET RECOVERY SIMPLE;   -- el DW se recarga desde el OLTP, no necesita log completo
GO
USE DW_Andina;
GO
-- Dimension TIEMPO (cuando)
CREATE TABLE dbo.DimTiempo (
    TiempoKey  INT PRIMARY KEY,             -- formato AAAAMMDD, ej. 20251224
    Fecha      DATE NOT NULL,
    Anio       SMALLINT NOT NULL,
    Trimestre  TINYINT NOT NULL,
    Mes        TINYINT NOT NULL,
    NombreMes  VARCHAR(12) NOT NULL,
    DiaSemana  VARCHAR(12) NOT NULL,
    EsCampana  BIT NOT NULL                 -- 1 = noviembre o diciembre
);
-- Dimension PRODUCTO (que)
CREATE TABLE dbo.DimProducto (
    ProductoKey INT IDENTITY(1,1) PRIMARY KEY,
    ProductoID  INT NOT NULL UNIQUE,        -- clave del sistema de origen
    Nombre      VARCHAR(100) NOT NULL,
    Categoria   VARCHAR(60) NOT NULL,
    PrecioLista DECIMAL(10,2) NOT NULL
);
-- Dimension CLIENTE (quien) - SIN datos personales (email, telefono, RUC)
CREATE TABLE dbo.DimCliente (
    ClienteKey  INT IDENTITY(1,1) PRIMARY KEY,
    ClienteID   INT NOT NULL UNIQUE,
    RazonSocial VARCHAR(120) NOT NULL,
    Segmento    VARCHAR(20) NOT NULL,
    Region      VARCHAR(50) NOT NULL,
    TieneEmail  CHAR(2) NOT NULL            -- 'SI' / 'NO' (regla R4)
);
-- Dimension SEDE (donde)
CREATE TABLE dbo.DimSede (
    SedeKey INT IDENTITY(1,1) PRIMARY KEY,
    SedeID  INT NOT NULL UNIQUE,
    Sede    VARCHAR(80) NOT NULL,
    Region  VARCHAR(50) NOT NULL
);
-- Tabla de HECHOS: una fila por producto vendido
CREATE TABLE dbo.FactVentas (
    TiempoKey      INT NOT NULL REFERENCES dbo.DimTiempo(TiempoKey),
    ProductoKey    INT NOT NULL REFERENCES dbo.DimProducto(ProductoKey),
    ClienteKey     INT NOT NULL REFERENCES dbo.DimCliente(ClienteKey),
    SedeKey        INT NOT NULL REFERENCES dbo.DimSede(SedeKey),
    VentaID        INT NOT NULL,            -- numero de venta original (trazabilidad)
    Cantidad       INT NOT NULL,
    PrecioUnitario DECIMAL(10,2) NOT NULL,
    Importe        DECIMAL(12,2) NOT NULL
);
-- Bitacora del ETL (cuantas filas entraron, se rechazaron y se cargaron)
CREATE TABLE dbo.ETL_Log (
    EjecucionID     INT IDENTITY(1,1) PRIMARY KEY,
    Inicio          DATETIME2(3) NOT NULL,
    Fin             DATETIME2(3) NULL,
    FilasExtraidas  INT NULL,
    FilasRechazadas INT NULL,
    FilasCargadas   INT NULL,
    Estado          VARCHAR(20) NULL,
    Mensaje         NVARCHAR(4000) NULL
);
GO

/* ---------- Procedimiento ETL ---------- */
CREATE OR ALTER PROCEDURE dbo.usp_ETL_CargarDW
AS
BEGIN
    SET NOCOUNT ON; SET XACT_ABORT ON;
    DECLARE @id INT, @ext INT, @rech INT, @carg INT;
    INSERT dbo.ETL_Log(Inicio, Estado) VALUES (SYSDATETIME(), 'EN PROCESO');
    SET @id = SCOPE_IDENTITY();
    BEGIN TRY
        BEGIN TRANSACTION;
        TRUNCATE TABLE dbo.FactVentas;                          -- recarga completa
        DELETE dbo.DimProducto; DELETE dbo.DimCliente; DELETE dbo.DimSede; DELETE dbo.DimTiempo;

        -- R3: atributos de tiempo derivados de la fecha (anio, trimestre, mes, campana)
        ;WITH f AS (SELECT CAST('2023-01-01' AS DATE) d UNION ALL SELECT DATEADD(DAY,1,d) FROM f WHERE d < '2026-12-31')
        INSERT dbo.DimTiempo
        SELECT CONVERT(INT, FORMAT(d,'yyyyMMdd')), d, YEAR(d), DATEPART(QUARTER,d), MONTH(d),
               UPPER(FORMAT(d,'MMMM','es-PE')), UPPER(FORMAT(d,'dddd','es-PE')),
               CASE WHEN MONTH(d) IN (11,12) THEN 1 ELSE 0 END
        FROM f OPTION (MAXRECURSION 2000);

        -- R1: estandarizar nombres (quitar espacios y pasar a MAYUSCULAS)
        INSERT dbo.DimProducto(ProductoID, Nombre, Categoria, PrecioLista)
        SELECT p.ProductoID, UPPER(LTRIM(RTRIM(p.Nombre))), c.Nombre, p.PrecioUnitario
        FROM DistribuidoraAndina.dbo.Producto p JOIN DistribuidoraAndina.dbo.Categoria c ON c.CategoriaID = p.CategoriaID;

        -- R4: minimizacion de datos personales -> no se copian email/telefono/RUC; solo un indicador
        INSERT dbo.DimCliente(ClienteID, RazonSocial, Segmento, Region, TieneEmail)
        SELECT c.ClienteID, UPPER(LTRIM(RTRIM(c.RazonSocial))), c.Segmento, r.Nombre,
               CASE WHEN NULLIF(LTRIM(c.Email),'') IS NULL THEN 'NO' ELSE 'SI' END
        FROM DistribuidoraAndina.dbo.Cliente c JOIN DistribuidoraAndina.dbo.Region r ON r.RegionID = c.RegionID;

        INSERT dbo.DimSede(SedeID, Sede, Region)
        SELECT s.SedeID, UPPER(s.Nombre), r.Nombre
        FROM DistribuidoraAndina.dbo.Sede s JOIN DistribuidoraAndina.dbo.Region r ON r.RegionID = s.RegionID;

        -- EXTRAER: todas las lineas de venta
        SELECT @ext = COUNT(*) FROM DistribuidoraAndina.dbo.DetalleVenta;

        -- R2: excluir ventas ANULADAS;  R5: rechazar cantidades/importes no validos
        INSERT dbo.FactVentas(TiempoKey, ProductoKey, ClienteKey, SedeKey, VentaID, Cantidad, PrecioUnitario, Importe)
        SELECT CONVERT(INT, FORMAT(v.Fecha,'yyyyMMdd')), dp.ProductoKey, dc.ClienteKey, ds.SedeKey,
               v.VentaID, d.Cantidad, d.PrecioUnitario, d.Subtotal
        FROM DistribuidoraAndina.dbo.DetalleVenta d
        JOIN DistribuidoraAndina.dbo.Venta v ON v.VentaID = d.VentaID
        JOIN dbo.DimProducto dp ON dp.ProductoID = d.ProductoID
        JOIN dbo.DimCliente  dc ON dc.ClienteID  = v.ClienteID
        JOIN dbo.DimSede     ds ON ds.SedeID     = v.SedeID
        WHERE v.Estado = 'EMITIDA'                  -- R2
          AND d.Cantidad > 0 AND d.Subtotal > 0;    -- R5
        SET @carg = @@ROWCOUNT;
        SET @rech = @ext - @carg;

        COMMIT TRANSACTION;
        UPDATE dbo.ETL_Log SET Fin = SYSDATETIME(), FilasExtraidas = @ext, FilasRechazadas = @rech,
               FilasCargadas = @carg, Estado = 'OK', Mensaje = 'Carga completa' WHERE EjecucionID = @id;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        UPDATE dbo.ETL_Log SET Fin = SYSDATETIME(), Estado = 'ERROR', Mensaje = ERROR_MESSAGE() WHERE EjecucionID = @id;
        THROW;
    END CATCH
    SELECT * FROM dbo.ETL_Log WHERE EjecucionID = @id;
END
GO
EXEC dbo.usp_ETL_CargarDW;
GO
-- Indice columnar: guarda la tabla de hechos "por columnas" y comprimida (ideal para sumar millones de filas)
CREATE CLUSTERED COLUMNSTORE INDEX CCI_FactVentas ON dbo.FactVentas;
GO
SELECT 'DimTiempo' Tabla, COUNT(*) Filas FROM dbo.DimTiempo UNION ALL
SELECT 'DimProducto', COUNT(*) FROM dbo.DimProducto UNION ALL
SELECT 'DimCliente', COUNT(*) FROM dbo.DimCliente UNION ALL
SELECT 'DimSede', COUNT(*) FROM dbo.DimSede UNION ALL
SELECT 'FactVentas', COUNT(*) FROM dbo.FactVentas;
GO

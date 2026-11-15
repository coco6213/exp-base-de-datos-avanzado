/* =====================================================================
   PROYECTO : DataShield 360 - Distribuidora Andina Comercial S.A.C.
   CURSO    : Base de Datos Avanzadas y Big Data (CIIN1487R)
   GRUPO    : Channel, Yin, Gutierrez, Pirsin
   SCRIPT 01: Crea la base transaccional (OLTP) con 11 tablas relacionadas.
   QUE HACE : Es la "base del dia a dia": aqui se guardan ventas, clientes,
              productos e inventario de las sedes Lima, Arequipa y Trujillo.
   ===================================================================== */
USE master;
GO
IF DB_ID('DistribuidoraAndina') IS NOT NULL
BEGIN
    ALTER DATABASE DistribuidoraAndina SET SINGLE_USER WITH ROLLBACK IMMEDIATE;
    DROP DATABASE DistribuidoraAndina;
END
GO
CREATE DATABASE DistribuidoraAndina;          -- base transaccional
GO
ALTER DATABASE DistribuidoraAndina SET RECOVERY FULL;  -- necesario para backup de log
GO
USE DistribuidoraAndina;
GO

-- 1. Regiones donde opera la empresa
CREATE TABLE dbo.Region (
    RegionID   INT IDENTITY(1,1) CONSTRAINT PK_Region PRIMARY KEY,
    Nombre     VARCHAR(50) NOT NULL CONSTRAINT UQ_Region_Nombre UNIQUE
);

-- 2. Sedes / almacenes (cada sede pertenece a una region)
CREATE TABLE dbo.Sede (
    SedeID     INT IDENTITY(1,1) CONSTRAINT PK_Sede PRIMARY KEY,
    Nombre     VARCHAR(80) NOT NULL,
    RegionID   INT NOT NULL CONSTRAINT FK_Sede_Region REFERENCES dbo.Region(RegionID)
);

-- 3. Categorias de producto
CREATE TABLE dbo.Categoria (
    CategoriaID INT IDENTITY(1,1) CONSTRAINT PK_Categoria PRIMARY KEY,
    Nombre      VARCHAR(60) NOT NULL CONSTRAINT UQ_Categoria_Nombre UNIQUE
);

-- 4. Productos de consumo masivo
CREATE TABLE dbo.Producto (
    ProductoID     INT IDENTITY(1,1) CONSTRAINT PK_Producto PRIMARY KEY,
    Nombre         VARCHAR(100) NOT NULL,
    CategoriaID    INT NOT NULL CONSTRAINT FK_Producto_Categoria REFERENCES dbo.Categoria(CategoriaID),
    PrecioUnitario DECIMAL(10,2) NOT NULL CONSTRAINT CK_Producto_Precio CHECK (PrecioUnitario > 0),
    Activo         BIT NOT NULL CONSTRAINT DF_Producto_Activo DEFAULT (1)
);

-- 5. Clientes (bodegas, minimarkets, mayoristas). Contiene datos personales -> tabla sensible
CREATE TABLE dbo.Cliente (
    ClienteID     INT IDENTITY(1,1) CONSTRAINT PK_Cliente PRIMARY KEY,
    RazonSocial   VARCHAR(120) NOT NULL,
    RUC           CHAR(11) NOT NULL CONSTRAINT UQ_Cliente_RUC UNIQUE,
    Email         VARCHAR(120) NULL,
    Telefono      VARCHAR(15) NULL,
    Segmento      VARCHAR(20) NOT NULL CONSTRAINT CK_Cliente_Segmento CHECK (Segmento IN ('BODEGA','MINIMARKET','MAYORISTA')),
    RegionID      INT NOT NULL CONSTRAINT FK_Cliente_Region REFERENCES dbo.Region(RegionID),
    FechaRegistro DATE NOT NULL CONSTRAINT DF_Cliente_Fecha DEFAULT (CAST(GETDATE() AS DATE))
);

-- 6. Vendedores asignados a una sede
CREATE TABLE dbo.Vendedor (
    VendedorID INT IDENTITY(1,1) CONSTRAINT PK_Vendedor PRIMARY KEY,
    Nombres    VARCHAR(100) NOT NULL,
    SedeID     INT NOT NULL CONSTRAINT FK_Vendedor_Sede REFERENCES dbo.Sede(SedeID)
);

-- 7. Inventario: stock de cada producto en cada sede
CREATE TABLE dbo.Inventario (
    SedeID      INT NOT NULL CONSTRAINT FK_Inv_Sede REFERENCES dbo.Sede(SedeID),
    ProductoID  INT NOT NULL CONSTRAINT FK_Inv_Producto REFERENCES dbo.Producto(ProductoID),
    Stock       INT NOT NULL CONSTRAINT CK_Inv_Stock CHECK (Stock >= 0),   -- nunca stock negativo
    StockMinimo INT NOT NULL CONSTRAINT DF_Inv_Min DEFAULT (20),
    CONSTRAINT PK_Inventario PRIMARY KEY (SedeID, ProductoID)
);

-- 8. Cabecera de venta
CREATE TABLE dbo.Venta (
    VentaID    INT IDENTITY(1,1) CONSTRAINT PK_Venta PRIMARY KEY,
    Fecha      DATETIME2(0) NOT NULL,
    ClienteID  INT NOT NULL CONSTRAINT FK_Venta_Cliente  REFERENCES dbo.Cliente(ClienteID),
    VendedorID INT NOT NULL CONSTRAINT FK_Venta_Vendedor REFERENCES dbo.Vendedor(VendedorID),
    SedeID     INT NOT NULL CONSTRAINT FK_Venta_Sede     REFERENCES dbo.Sede(SedeID),
    Total      DECIMAL(12,2) NOT NULL CONSTRAINT DF_Venta_Total DEFAULT (0),
    Estado     VARCHAR(10) NOT NULL CONSTRAINT CK_Venta_Estado CHECK (Estado IN ('EMITIDA','ANULADA'))
);

-- 9. Detalle de venta (productos de cada venta)
CREATE TABLE dbo.DetalleVenta (
    DetalleID      INT IDENTITY(1,1) CONSTRAINT PK_DetalleVenta PRIMARY KEY,
    VentaID        INT NOT NULL CONSTRAINT FK_Det_Venta    REFERENCES dbo.Venta(VentaID),
    ProductoID     INT NOT NULL CONSTRAINT FK_Det_Producto REFERENCES dbo.Producto(ProductoID),
    Cantidad       INT NOT NULL CONSTRAINT CK_Det_Cantidad CHECK (Cantidad > 0),
    PrecioUnitario DECIMAL(10,2) NOT NULL,
    Subtotal       AS (CAST(Cantidad * PrecioUnitario AS DECIMAL(12,2))) PERSISTED
);

-- 10. Auditoria: lo llenan los triggers (quien cambio que y cuando)
CREATE TABLE dbo.AuditoriaLog (
    AuditID         BIGINT IDENTITY(1,1) CONSTRAINT PK_AuditoriaLog PRIMARY KEY,
    Tabla           SYSNAME      NOT NULL,
    Operacion       CHAR(6)      NOT NULL,           -- INSERT / UPDATE / DELETE
    Usuario         SYSNAME      NOT NULL CONSTRAINT DF_Aud_Usuario DEFAULT (SUSER_SNAME()),
    Fecha           DATETIME2(3) NOT NULL CONSTRAINT DF_Aud_Fecha DEFAULT (SYSDATETIME()),
    DatosAnteriores NVARCHAR(MAX) NULL,              -- JSON con la fila antes del cambio
    DatosNuevos     NVARCHAR(MAX) NULL               -- JSON con la fila despues del cambio
);

-- 11. Alertas de stock bajo (las genera un procedimiento)
CREATE TABLE dbo.AlertaStock (
    AlertaID   INT IDENTITY(1,1) CONSTRAINT PK_AlertaStock PRIMARY KEY,
    SedeID     INT NOT NULL,
    ProductoID INT NOT NULL,
    Stock      INT NOT NULL,
    Fecha      DATETIME2(0) NOT NULL CONSTRAINT DF_Alerta_Fecha DEFAULT (SYSDATETIME()),
    CONSTRAINT FK_Alerta_Inv FOREIGN KEY (SedeID, ProductoID) REFERENCES dbo.Inventario(SedeID, ProductoID)
);
GO
PRINT 'Base DistribuidoraAndina creada con 11 tablas.';

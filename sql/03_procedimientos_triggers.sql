/* =====================================================================
   SCRIPT 03: Automatizacion - 3 procedimientos almacenados + 2 triggers.
   EN SIMPLE: un PROCEDIMIENTO es una "receta guardada" que la base ejecuta
   siempre igual; un TRIGGER es un "vigilante" que se activa solo cuando
   alguien cambia una tabla y deja anotado el cambio.
   TRANSACCION: o se guarda TODO (COMMIT) o no se guarda NADA (ROLLBACK).
   ===================================================================== */
USE DistribuidoraAndina;
GO

/* ---------- SP 1: genera alertas de stock bajo ---------- */
CREATE OR ALTER PROCEDURE dbo.usp_GenerarAlertasStock
    @SedeID INT = NULL                         -- NULL = todas las sedes
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        INSERT dbo.AlertaStock(SedeID, ProductoID, Stock)
        SELECT i.SedeID, i.ProductoID, i.Stock
        FROM dbo.Inventario i
        WHERE i.Stock < i.StockMinimo                        -- stock por debajo del minimo
          AND (@SedeID IS NULL OR i.SedeID = @SedeID)
          AND NOT EXISTS (SELECT 1 FROM dbo.AlertaStock a    -- evita alertas repetidas el mismo dia
                          WHERE a.SedeID = i.SedeID AND a.ProductoID = i.ProductoID
                            AND CAST(a.Fecha AS DATE) = CAST(SYSDATETIME() AS DATE));
        SELECT @@ROWCOUNT AS AlertasGeneradas;
    END TRY
    BEGIN CATCH
        THROW;                                               -- devuelve el error a quien llamo
    END CATCH
END
GO

/* ---------- SP 2: registra una venta completa (cabecera + detalle + stock) ---------- */
CREATE OR ALTER PROCEDURE dbo.usp_RegistrarVenta
    @ClienteID  INT,
    @VendedorID INT,
    @SedeID     INT,
    @Detalle    NVARCHAR(MAX)    -- JSON: [{"ProductoID":1,"Cantidad":3}, ...]
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;                                       -- ante error grave, deshace todo
    DECLARE @VentaID INT;
    BEGIN TRY
        BEGIN TRANSACTION;                                   -- inicio de la transaccion

        INSERT dbo.Venta(Fecha, ClienteID, VendedorID, SedeID, Estado)
        VALUES (SYSDATETIME(), @ClienteID, @VendedorID, @SedeID, 'EMITIDA');
        SET @VentaID = SCOPE_IDENTITY();

        INSERT dbo.DetalleVenta(VentaID, ProductoID, Cantidad, PrecioUnitario)
        SELECT @VentaID, j.ProductoID, j.Cantidad, p.PrecioUnitario
        FROM OPENJSON(@Detalle) WITH (ProductoID INT, Cantidad INT) j
        JOIN dbo.Producto p ON p.ProductoID = j.ProductoID;

        -- descuenta stock; si queda negativo la regla CHECK falla -> salta al CATCH
        UPDATE i SET i.Stock = i.Stock - j.Cantidad
        FROM dbo.Inventario i
        JOIN OPENJSON(@Detalle) WITH (ProductoID INT, Cantidad INT) j ON j.ProductoID = i.ProductoID
        WHERE i.SedeID = @SedeID;

        UPDATE dbo.Venta SET Total = (SELECT SUM(Subtotal) FROM dbo.DetalleVenta WHERE VentaID = @VentaID)
        WHERE VentaID = @VentaID;

        COMMIT TRANSACTION;                                  -- todo OK: se guarda
        EXEC dbo.usp_GenerarAlertasStock @SedeID;            -- revisa si algo quedo con poco stock
        SELECT @VentaID AS VentaRegistrada, 'OK' AS Resultado;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;             -- algo fallo: no se guarda nada
        SELECT NULL AS VentaRegistrada, 'ERROR: ' + ERROR_MESSAGE() AS Resultado;
    END CATCH
END
GO

/* ---------- SP 3: repone / ajusta inventario ---------- */
CREATE OR ALTER PROCEDURE dbo.usp_ActualizarInventario
    @SedeID INT, @ProductoID INT, @Cantidad INT              -- positivo = ingreso, negativo = salida
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        BEGIN TRANSACTION;
        UPDATE dbo.Inventario SET Stock = Stock + @Cantidad
        WHERE SedeID = @SedeID AND ProductoID = @ProductoID;
        IF @@ROWCOUNT = 0
            THROW 50001, 'El producto no existe en el inventario de esa sede.', 1;
        COMMIT TRANSACTION;
        SELECT SedeID, ProductoID, Stock, 'OK' AS Resultado FROM dbo.Inventario
        WHERE SedeID = @SedeID AND ProductoID = @ProductoID;
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        SELECT @SedeID SedeID, @ProductoID ProductoID, NULL Stock, 'ERROR: ' + ERROR_MESSAGE() AS Resultado;
    END CATCH
END
GO

/* ---------- TRIGGER 1: auditoria de la tabla Cliente (datos personales) ---------- */
CREATE OR ALTER TRIGGER dbo.trg_Auditoria_Cliente
ON dbo.Cliente AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @op CHAR(6) = CASE WHEN EXISTS(SELECT 1 FROM inserted) AND EXISTS(SELECT 1 FROM deleted) THEN 'UPDATE'
                                   WHEN EXISTS(SELECT 1 FROM inserted) THEN 'INSERT' ELSE 'DELETE' END;
        INSERT dbo.AuditoriaLog(Tabla, Operacion, DatosAnteriores, DatosNuevos)
        SELECT 'Cliente', @op,
               (SELECT * FROM deleted  FOR JSON PATH),       -- como estaba antes
               (SELECT * FROM inserted FOR JSON PATH);       -- como quedo despues
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;             -- si no se puede auditar, no se permite el cambio
        THROW;
    END CATCH
END
GO

/* ---------- TRIGGER 2: auditoria de la tabla Producto (cambios de precio) ---------- */
CREATE OR ALTER TRIGGER dbo.trg_Auditoria_Producto
ON dbo.Producto AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @op CHAR(6) = CASE WHEN EXISTS(SELECT 1 FROM inserted) AND EXISTS(SELECT 1 FROM deleted) THEN 'UPDATE'
                                   WHEN EXISTS(SELECT 1 FROM inserted) THEN 'INSERT' ELSE 'DELETE' END;
        INSERT dbo.AuditoriaLog(Tabla, Operacion, DatosAnteriores, DatosNuevos)
        SELECT 'Producto', @op,
               (SELECT * FROM deleted  FOR JSON PATH),
               (SELECT * FROM inserted FOR JSON PATH);
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
        THROW;
    END CATCH
END
GO
PRINT 'Procedimientos y triggers creados.';

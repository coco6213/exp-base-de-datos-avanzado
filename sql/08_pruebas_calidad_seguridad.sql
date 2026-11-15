/* =====================================================================
   SCRIPT 08: Pruebas de calidad, automatizacion y seguridad (T-01..T-08).
   EN SIMPLE: comprobamos que cada pieza hace lo que promete:
   procedimientos, transacciones, triggers de auditoria, roles y ETL.
   ===================================================================== */
USE DistribuidoraAndina;
GO
SET NOCOUNT ON;

PRINT '--- T-01: Registrar una venta valida (debe responder OK) ---';
EXEC dbo.usp_RegistrarVenta @ClienteID = 10, @VendedorID = 1, @SedeID = 1,
     @Detalle = N'[{"ProductoID":5,"Cantidad":2},{"ProductoID":7,"Cantidad":1}]';

PRINT '--- T-02: Venta con stock insuficiente (debe hacer ROLLBACK y no guardar nada) ---';
DECLARE @antes INT = (SELECT COUNT(*) FROM dbo.Venta);
EXEC dbo.usp_RegistrarVenta @ClienteID = 10, @VendedorID = 1, @SedeID = 1,
     @Detalle = N'[{"ProductoID":5,"Cantidad":999999}]';
SELECT @antes AS VentasAntes, (SELECT COUNT(*) FROM dbo.Venta) AS VentasDespues,
       CASE WHEN @antes = (SELECT COUNT(*) FROM dbo.Venta) THEN 'PASA: ROLLBACK correcto' ELSE 'FALLA' END AS T02;

PRINT '--- T-03: Reponer inventario y generar alerta de stock bajo ---';
EXEC dbo.usp_ActualizarInventario @SedeID = 2, @ProductoID = 3, @Cantidad = -1 ;   -- salida pequena
UPDATE dbo.Inventario SET Stock = 5 WHERE SedeID = 2 AND ProductoID = 3;           -- simulamos stock critico
EXEC dbo.usp_GenerarAlertasStock @SedeID = 2;
SELECT TOP 3 * FROM dbo.AlertaStock ORDER BY AlertaID DESC;
EXEC dbo.usp_ActualizarInventario @SedeID = 99, @ProductoID = 3, @Cantidad = 10;   -- sede inexistente -> error controlado

PRINT '--- T-04: Trigger de auditoria en Cliente (UPDATE de email) ---';
UPDATE dbo.Cliente SET Email = 'nuevo.correo@correo.pe' WHERE ClienteID = 100;
SELECT TOP 1 AuditID, Tabla, Operacion, Usuario, Fecha,
       LEFT(DatosAnteriores, 120) AS Antes, LEFT(DatosNuevos, 120) AS Despues
FROM dbo.AuditoriaLog WHERE Tabla = 'Cliente' ORDER BY AuditID DESC;

PRINT '--- T-05: Trigger de auditoria en Producto (cambio de precio) ---';
UPDATE dbo.Producto SET PrecioUnitario = PrecioUnitario * 1.05 WHERE ProductoID = 1;
SELECT TOP 1 AuditID, Tabla, Operacion, Usuario, Fecha,
       JSON_VALUE(DatosAnteriores, '$[0].PrecioUnitario') AS PrecioAntes,
       JSON_VALUE(DatosNuevos, '$[0].PrecioUnitario') AS PrecioDespues
FROM dbo.AuditoriaLog WHERE Tabla = 'Producto' ORDER BY AuditID DESC;

PRINT '--- T-06: Rol analista NO puede ver datos personales ---';
EXECUTE AS USER = 'usr_analista';
BEGIN TRY
    SELECT TOP 1 ClienteID, RazonSocial, Segmento FROM dbo.Cliente;          -- permitido
    SELECT TOP 1 Email FROM dbo.Cliente;                                      -- debe fallar
END TRY
BEGIN CATCH
    SELECT 'PASA: acceso denegado' AS T06, ERROR_MESSAGE() AS Mensaje;
END CATCH
REVERT;

PRINT '--- T-07: Rol operativo NO puede borrar ventas directamente ---';
EXECUTE AS USER = 'usr_operativo';
BEGIN TRY
    DELETE FROM dbo.Venta WHERE VentaID = 1;
END TRY
BEGIN CATCH
    SELECT 'PASA: DELETE bloqueado' AS T07, ERROR_MESSAGE() AS Mensaje;
END CATCH
REVERT;

GO


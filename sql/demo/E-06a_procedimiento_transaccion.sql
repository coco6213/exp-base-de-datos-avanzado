-- E-06a | Procedimiento usp_RegistrarVenta: venta OK y venta con ROLLBACK
USE DistribuidoraAndina;
EXEC dbo.usp_RegistrarVenta @ClienteID=20, @VendedorID=1, @SedeID=1, @Detalle=N'[{"ProductoID":8,"Cantidad":2}]';      -- valida -> OK
EXEC dbo.usp_RegistrarVenta @ClienteID=20, @VendedorID=1, @SedeID=1, @Detalle=N'[{"ProductoID":8,"Cantidad":999999}]'; -- sin stock -> ROLLBACK
SELECT COUNT(*) AS TotalVentas FROM dbo.Venta;   -- solo aumento en 1 (la venta fallida no se guardo)

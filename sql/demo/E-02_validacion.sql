-- E-02 | Validacion del dataset (perfil de calidad) - DistribuidoraAndina
USE DistribuidoraAndina;
SELECT 'Venta' Tabla, COUNT(*) Filas FROM dbo.Venta UNION ALL SELECT 'DetalleVenta', COUNT(*) FROM dbo.DetalleVenta
UNION ALL SELECT 'Cliente', COUNT(*) FROM dbo.Cliente UNION ALL SELECT 'Producto', COUNT(*) FROM dbo.Producto;
SELECT 'V2 Ventas sin detalle' Prueba, COUNT(*) Problemas FROM dbo.Venta v WHERE NOT EXISTS (SELECT 1 FROM dbo.DetalleVenta d WHERE d.VentaID=v.VentaID)
UNION ALL SELECT 'V4 RUC duplicados', COUNT(*)-COUNT(DISTINCT RUC) FROM dbo.Cliente
UNION ALL SELECT 'V5 Total <> suma detalle', COUNT(*) FROM dbo.Venta v JOIN (SELECT VentaID, SUM(Subtotal) s FROM dbo.DetalleVenta GROUP BY VentaID) d ON d.VentaID=v.VentaID WHERE v.Total<>d.s
UNION ALL SELECT 'V9 Clientes sin email (dato sucio)', COUNT(*) FROM dbo.Cliente WHERE Email IS NULL
UNION ALL SELECT 'V11 Ventas anuladas (dato a excluir)', COUNT(*) FROM dbo.Venta WHERE Estado='ANULADA';

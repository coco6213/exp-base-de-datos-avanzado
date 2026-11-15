/* =====================================================================
   SCRIPT 11: Reconciliacion ETL (T-08). Se ejecuta DESPUES del ETL (script 06).
   EN SIMPLE: el total vendido en la base diaria (sin anuladas) debe ser
   igual, al centimo, al total del Data Warehouse.
   ===================================================================== */
USE DistribuidoraAndina;
GO
PRINT '--- T-08: Reconciliacion ETL (origen vs Data Warehouse) ---';
SELECT o.Importe AS ImporteOLTP, d.Importe AS ImporteDW,
       CASE WHEN o.Importe = d.Importe THEN 'PASA: cuadra al centimo' ELSE 'FALLA' END AS T08
FROM (SELECT SUM(dv.Subtotal) Importe FROM dbo.DetalleVenta dv JOIN dbo.Venta v ON v.VentaID = dv.VentaID
      WHERE v.Estado = 'EMITIDA') o
CROSS JOIN (SELECT SUM(Importe) Importe FROM DW_Andina.dbo.FactVentas) d;
GO


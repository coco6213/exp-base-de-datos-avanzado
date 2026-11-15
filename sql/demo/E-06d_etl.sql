-- E-06d | ETL: bitacora de carga y reconciliacion OLTP vs Data Warehouse
USE DW_Andina;
SELECT EjecucionID, Inicio, Fin, FilasExtraidas, FilasRechazadas, FilasCargadas, Estado FROM dbo.ETL_Log;
SELECT (SELECT SUM(dv.Subtotal) FROM DistribuidoraAndina.dbo.DetalleVenta dv JOIN DistribuidoraAndina.dbo.Venta v ON v.VentaID=dv.VentaID
        WHERE v.Estado='EMITIDA' AND v.VentaID <= 200001) AS ImporteOLTP,
       (SELECT SUM(Importe) FROM dbo.FactVentas) AS ImporteDW;

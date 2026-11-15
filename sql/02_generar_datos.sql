/* =====================================================================
   SCRIPT 02: Genera el dataset simulado (reproducible).
   QUE HACE : Llena la base con datos ficticios pero realistas:
              3 regiones, 6 sedes, 8 categorias, 200 productos, 5 000 clientes,
              30 vendedores, 200 000 ventas y ~600 000 lineas de detalle (2023-2025).
   REPRODUCIBLE: no usa NEWID(); los valores salen de un "hash" del numero de fila,
              por eso cada vez que se ejecuta se obtiene EXACTAMENTE el mismo dataset.
   Tambien mete "datos sucios" a proposito (espacios, minusculas, emails vacios,
   ventas anuladas) para probar las reglas de limpieza del ETL.
   ===================================================================== */
USE DistribuidoraAndina;
GO
SET NOCOUNT ON;

-- Tabla auxiliar de numeros 1..1 000 000 (sirve para generar filas en bloque)
DROP TABLE IF EXISTS #N;
SELECT TOP (1000000) n = ROW_NUMBER() OVER (ORDER BY (SELECT NULL))
INTO #N FROM sys.all_objects a CROSS JOIN sys.all_objects b CROSS JOIN sys.all_objects c;
CREATE UNIQUE CLUSTERED INDEX ix ON #N(n);

-- Funcion "aleatoria" determinista: h(n, semilla) -> entero positivo
-- CAST(SUBSTRING(HASHBYTES('MD5', ...),1,4) AS BIGINT) devuelve 0..4 294 967 295
DECLARE @inicio DATE = '2023-01-01';

INSERT dbo.Region(Nombre) VALUES ('LIMA'),('AREQUIPA'),('TRUJILLO');

INSERT dbo.Sede(Nombre, RegionID) VALUES
 ('Almacen Lima Norte',1),('Almacen Lima Sur',1),
 ('Almacen Arequipa Centro',2),('Almacen Cayma',2),
 ('Almacen Trujillo Centro',3),('Almacen La Esperanza',3);

INSERT dbo.Categoria(Nombre) VALUES
 ('ABARROTES'),('BEBIDAS'),('LACTEOS'),('LIMPIEZA'),('CUIDADO PERSONAL'),('SNACKS'),('CONSERVAS'),('PANADERIA');

-- 200 productos; 1 de cada 25 con nombre "sucio" (espacios + minusculas)
INSERT dbo.Producto(Nombre, CategoriaID, PrecioUnitario)
SELECT CASE WHEN n % 25 = 0 THEN '  producto ' + RIGHT('000'+CAST(n AS VARCHAR),3) + '  '
            ELSE 'PRODUCTO ' + RIGHT('000'+CAST(n AS VARCHAR),3) END,
       (n % 8) + 1,
       CAST(1.50 + (CAST(SUBSTRING(HASHBYTES('MD5','p'+CAST(n AS VARCHAR)),1,4) AS BIGINT) % 4850) / 100.0 AS DECIMAL(10,2))
FROM #N WHERE n <= 200;

-- 5 000 clientes; 1 de cada 100 sin email (dato faltante)
INSERT dbo.Cliente(RazonSocial, RUC, Email, Telefono, Segmento, RegionID, FechaRegistro)
SELECT 'CLIENTE ' + RIGHT('0000'+CAST(n AS VARCHAR),5) + ' S.A.C.',
       '20' + RIGHT('000000000' + CAST(100000000 + n AS VARCHAR), 9),
       CASE WHEN n % 100 = 0 THEN NULL ELSE 'cliente' + CAST(n AS VARCHAR) + '@correo.pe' END,
       '9' + RIGHT('00000000' + CAST(n * 7 AS VARCHAR), 8),
       -- v2: el segmento usa (n/10)%10 para que sea INDEPENDIENTE de la region (que usa n%10).
       -- En la v1 ambos usaban n%10 y todos los clientes de Lima salian como BODEGA (sesgo detectado en el dashboard).
       CASE WHEN (n / 10) % 10 < 6 THEN 'BODEGA' WHEN (n / 10) % 10 < 9 THEN 'MINIMARKET' ELSE 'MAYORISTA' END,
       CASE WHEN n % 10 < 5 THEN 1 WHEN n % 10 < 8 THEN 2 ELSE 3 END,     -- 50% Lima, 30% Arequipa, 20% Trujillo
       DATEADD(DAY, -(n % 900), '2023-01-01')
FROM #N WHERE n <= 5000;

-- 30 vendedores (5 por sede)
INSERT dbo.Vendedor(Nombres, SedeID)
SELECT 'VENDEDOR ' + RIGHT('00'+CAST(n AS VARCHAR),2), ((n-1) % 6) + 1 FROM #N WHERE n <= 30;

-- Inventario inicial de cada producto en cada sede
INSERT dbo.Inventario(SedeID, ProductoID, Stock, StockMinimo)
SELECT s.SedeID, p.ProductoID,
       50 + CAST(SUBSTRING(HASHBYTES('MD5','i'+CAST(s.SedeID*1000+p.ProductoID AS VARCHAR)),1,4) AS BIGINT) % 950,
       20
FROM dbo.Sede s CROSS JOIN dbo.Producto p;

-- 200 000 ventas entre 2023-01-01 y 2025-12-31.
-- Estacionalidad: 20% de las ventas cae en noviembre-diciembre (campana navidena).
-- 3% de las ventas queda ANULADA (el ETL debe excluirlas).
;WITH v AS (
    SELECT n, h = CAST(SUBSTRING(HASHBYTES('MD5','v'+CAST(n AS VARCHAR)),1,4) AS BIGINT),
              h2 = CAST(SUBSTRING(HASHBYTES('MD5','w'+CAST(n AS VARCHAR)),1,4) AS BIGINT)
    FROM #N WHERE n <= 200000
)
INSERT dbo.Venta(Fecha, ClienteID, VendedorID, SedeID, Estado)
SELECT  DATEADD(SECOND, CAST(h2 % 36000 AS INT) + 28800,                -- hora entre 08:00 y 18:00
          CAST(CASE WHEN h % 10 < 2   -- campana Nov-Dic
                    THEN DATEADD(DAY, CAST(h2 % 61 AS INT), DATEFROMPARTS(2023 + CAST(h % 3 AS INT), 11, 1))
                    ELSE DATEADD(DAY, CAST(h % 1096 AS INT), @inicio) END AS DATETIME2(0))),
        c.ClienteID,
        vd.VendedorID,
        vd.SedeID,
        CASE WHEN h2 % 100 < 3 THEN 'ANULADA' ELSE 'EMITIDA' END
FROM v
CROSS APPLY (SELECT ClienteID = CAST(h % 5000 AS INT) + 1) c
CROSS APPLY (SELECT TOP 1 VendedorID, SedeID FROM dbo.Vendedor
             WHERE SedeID IN (SELECT SedeID FROM dbo.Sede s JOIN dbo.Cliente cl ON cl.RegionID = s.RegionID AND cl.ClienteID = c.ClienteID)
             ORDER BY (VendedorID * 31 + v.n) % 97) vd           -- vendedor de la region del cliente
ORDER BY v.n;                                                    -- garantiza que VentaID = n (reproducible)

-- Detalle: entre 1 y 5 productos por venta (~600 000 lineas)
;WITH lineas AS (
    SELECT v.VentaID, k.k
    FROM dbo.Venta v
    CROSS APPLY (SELECT TOP (1 + CAST(SUBSTRING(HASHBYTES('MD5','d'+CAST(v.VentaID AS VARCHAR)),1,4) AS BIGINT) % 5) k = n FROM #N ORDER BY n) k
)
INSERT dbo.DetalleVenta(VentaID, ProductoID, Cantidad, PrecioUnitario)
SELECT l.VentaID, p.ProductoID, q.Cantidad, p.PrecioUnitario
FROM lineas l
CROSS APPLY (SELECT ProdID = 1 + CAST(SUBSTRING(HASHBYTES('MD5','x'+CAST(l.VentaID*10+l.k AS VARCHAR)),1,4) AS BIGINT) % 200) r
CROSS APPLY (SELECT Cantidad = 1 + CAST(SUBSTRING(HASHBYTES('MD5','q'+CAST(l.VentaID*10+l.k AS VARCHAR)),1,4) AS BIGINT) % 24) q
JOIN dbo.Producto p ON p.ProductoID = r.ProdID
ORDER BY l.VentaID, l.k;                                         -- orden fijo de DetalleID

-- Total de cada venta = suma de su detalle
UPDATE v SET Total = d.Total
FROM dbo.Venta v JOIN (SELECT VentaID, SUM(Subtotal) Total FROM dbo.DetalleVenta GROUP BY VentaID) d ON d.VentaID = v.VentaID;

SELECT 'Region' Tabla, COUNT(*) Filas FROM dbo.Region UNION ALL
SELECT 'Sede', COUNT(*) FROM dbo.Sede UNION ALL
SELECT 'Categoria', COUNT(*) FROM dbo.Categoria UNION ALL
SELECT 'Producto', COUNT(*) FROM dbo.Producto UNION ALL
SELECT 'Cliente', COUNT(*) FROM dbo.Cliente UNION ALL
SELECT 'Vendedor', COUNT(*) FROM dbo.Vendedor UNION ALL
SELECT 'Inventario', COUNT(*) FROM dbo.Inventario UNION ALL
SELECT 'Venta', COUNT(*) FROM dbo.Venta UNION ALL
SELECT 'DetalleVenta', COUNT(*) FROM dbo.DetalleVenta;
GO

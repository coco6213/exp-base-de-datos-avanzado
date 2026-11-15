# =====================================================================
# Crea el dashboard BI en Excel conectado al Data Warehouse (DW_Andina).
# Hojas: Dashboard (KPIs + 4 graficos + segmentadores), Tablas_OLAP, Datos.
# Uso:  powershell -File crear_dashboard_excel.ps1
# =====================================================================
param([string]$Salida = "$PSScriptRoot\..\dashboard\Dashboard_Andina.xlsx")
$ErrorActionPreference = 'Stop'
$Salida = [IO.Path]::GetFullPath($Salida)
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null

# Abre Microsoft Excel (no el visor por defecto) y se conecta por automatizacion
$exe = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { Start-Process $exe -ArgumentList '/automation','-Embedding'; Start-Sleep 6 }
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -lt 3) { $wb.Worksheets.Add() | Out-Null }
$dash = $wb.Worksheets.Item(1); $dash.Name = 'Dashboard'
$olap = $wb.Worksheets.Item(2); $olap.Name = 'Tablas_OLAP'
$dat  = $wb.Worksheets.Item(3); $dat.Name  = 'Datos'

# 1) Conexion ODBC al Data Warehouse (se puede actualizar con Datos > Actualizar todo)
$con = 'ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost;DATABASE=DW_Andina;Trusted_Connection=Yes;TrustServerCertificate=Yes;'
$lo = $dat.ListObjects.Add(0, $con, $true, 1, $dat.Range('A1'))
$lo.QueryTable.CommandText = 'SELECT Anio, Mes, Periodo, Trimestre, Temporada, Region, Sede, Categoria, Producto, Segmento, Unidades, Importe, Ventas FROM dbo.vw_Dashboard'
$lo.QueryTable.Refresh($false) | Out-Null
$lo.Name = 'tblVentas'
$lo.ListColumns.Item('Importe').DataBodyRange.NumberFormat = '#,##0.00'

# 2) Tablas dinamicas (cubo OLAP) sobre la misma cache
$pc = $wb.PivotCaches().Create(1, 'tblVentas')
function NuevaPivot($nombre, $celda) { $pc.CreatePivotTable($olap.Range($celda), $nombre) }
$xlRow=1; $xlCol=2; $xlSum=-4157

# P0 KPIs
$p0 = NuevaPivot 'ptKPI' 'A3'
$f=$p0.AddDataField($p0.PivotFields('Importe'), 'Ventas S/', $xlSum); $f.NumberFormat='#,##0'
$f=$p0.AddDataField($p0.PivotFields('Unidades'), 'Unidades vendidas', $xlSum); $f.NumberFormat='#,##0'
# P1 Tendencia mensual
$p1 = NuevaPivot 'ptTendencia' 'E3'
$p1.PivotFields('Periodo').Orientation = $xlRow
$f=$p1.AddDataField($p1.PivotFields('Importe'), 'Ventas mensuales S/', $xlSum); $f.NumberFormat='#,##0'
# P2 Region x Anio (OLAP tiempo x region)
$p2 = NuevaPivot 'ptRegionAnio' 'I3'
$p2.PivotFields('Region').Orientation = $xlRow
$p2.PivotFields('Anio').Orientation = $xlCol
$f=$p2.AddDataField($p2.PivotFields('Importe'), 'Ventas region', $xlSum); $f.NumberFormat='#,##0'
# P3 Top 10 productos
$p3 = NuevaPivot 'ptTopProductos' 'P3'
$pf = $p3.PivotFields('Producto'); $pf.Orientation = $xlRow
$f=$p3.AddDataField($p3.PivotFields('Importe'), 'Ventas producto', $xlSum); $f.NumberFormat='#,##0'
$pf.PivotFilters.Add2(1, $p3.DataFields(1), 10) | Out-Null   # filtro Top 10 (xlTopCount)
$pf.AutoSort(1, $p3.DataFields(1).Name)              # ascendente -> barra mayor arriba
# P4 Categoria
$p4 = NuevaPivot 'ptCategoria' 'T3'
$pf = $p4.PivotFields('Categoria'); $pf.Orientation = $xlRow
$f=$p4.AddDataField($p4.PivotFields('Importe'), 'Ventas categoria', $xlSum); $f.NumberFormat='#,##0'
$pf.AutoSort(2, $p4.DataFields(1).Name)
$olap.Range('A1').Value2 = 'Tablas OLAP (tablas dinamicas) que alimentan el dashboard - fuente: DW_Andina.dbo.vw_Dashboard'
$olap.Range('A1').Font.Bold = $true

# 3) Hoja Dashboard
$xl.ActiveWindow.DisplayGridlines = $true
$dash.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
$dash.Cells.Interior.Color = 0xF7F3F0          # fondo gris claro (BGR)
$dash.Range('A1:Z1').RowHeight = 34
$t = $dash.Range('B1'); $t.Value2 = 'Distribuidora Andina Comercial S.A.C. - Dashboard de Ventas 2023-2025'
$t.Font.Size = 20; $t.Font.Bold = $true; $t.Font.Color = 0x5A3A1F
$dash.Range('B2').Value2 = 'Fuente: Data Warehouse DW_Andina (esquema estrella). Filtre con los segmentadores de la izquierda. Grupo: Channel, Yin, Gutierrez, Pirsin.'
$dash.Range('B2').Font.Color = 0x666666
$dash.Columns('A').ColumnWidth = 2

# KPI tiles (formulas GETPIVOTDATA -> responden a los filtros)
function Tile($rango, $titulo, $formula, $fmt) {
  $r = $dash.Range($rango); $r.Merge(); $r.Interior.Color = 0xFFFFFF
  $r.Borders.Color = 0xD9D9D9
  $c = $r.Cells.Item(1,1)
  $c.Formula = $formula; $c.NumberFormat = $fmt
  $c.Font.Size = 22; $c.Font.Bold = $true; $c.Font.Color = 0x8B5A1F
  $c.HorizontalAlignment = -4108; $c.VerticalAlignment = -4108
  $h = $r.Offset(-1,0).Resize(1, $r.Columns.Count); $h.Merge(); $h.Value2 = $titulo
  $h.Font.Color = 0x555555; $h.Font.Bold = $true; $h.HorizontalAlignment = -4108
}
Tile 'F4:H5' 'VENTAS TOTALES (S/)'     '=GETPIVOTDATA("Ventas S/",Tablas_OLAP!$A$3)' '"S/ "#,##0'
Tile 'J4:L5' 'UNIDADES VENDIDAS'        '=GETPIVOTDATA("Unidades vendidas",Tablas_OLAP!$A$3)'  '#,##0'
Tile 'N4:P5' 'PRECIO PROMEDIO / UNIDAD' '=GETPIVOTDATA("Ventas S/",Tablas_OLAP!$A$3)/GETPIVOTDATA("Unidades vendidas",Tablas_OLAP!$A$3)' '"S/ "0.00'
Tile 'R4:T5' 'PRODUCTOS EN EL TOP 10'   '=COUNTA(Tablas_OLAP!P5:P14)' '0'

# Graficos
function Grafico($pt, $tipo, $titulo, $left, $top, $w, $h) {
  $sh = $dash.Shapes.AddChart2(-1, $tipo, $left, $top, $w, $h)
  $sh.Chart.SetSourceData($pt.TableRange1)
  $sh.Chart.HasTitle = $true; $sh.Chart.ChartTitle.Text = $titulo
  $sh.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
  try { $sh.Chart.ShowAllFieldButtons = $false } catch {}
  return $sh.Chart
}
$L = $dash.Range('F7').Left; $T = $dash.Range('F7').Top
$c1 = Grafico $p1 4  'Tendencia mensual de ventas (S/) - picos en campana Nov-Dic' $L $T 520 250
$c1.HasLegend = $false
$c2 = Grafico $p2 51 'Ventas por region y anio (S/) - analisis OLAP' ($L+530) $T 420 250
$c3 = Grafico $p3 57 'Top 10 productos por ventas (S/)' $L ($T+260) 520 260
$c3.HasLegend = $false
$c3.Axes(2).MinimumScale = 0      # eje desde cero: no exagera diferencias
$c4 = Grafico $p4 -4120 'Participacion por categoria' ($L+530) ($T+260) 420 260
$c4.ApplyDataLabels(3)   # porcentaje
$c2.HasLegend = $true; $c2.Legend.Position = -4107     # leyenda de anios abajo
$c4.HasLegend = $true; $c4.Legend.Position = -4152     # leyenda de categorias a la derecha

# Segmentadores (filtran KPIs y los 4 graficos a la vez)
$top = $dash.Range('B4').Top; $i = 0
foreach ($campo in 'Anio','Region','Segmento','Temporada') {
  $sc = $wb.SlicerCaches.Add2($p0, $campo)
  foreach ($p in $p1,$p2,$p3,$p4) { $sc.PivotTables.AddPivotTable($p) }
  $alto = @{Anio=110;Region=110;Segmento=110;Temporada=90}[$campo]
  $sc.Slicers.Add($dash, [Type]::Missing, "sl$campo", $campo, $top, 6, 150, $alto) | Out-Null
  $top += $alto + 8
}
$dash.Columns('B:D').ColumnWidth = 8
$dash.Range('A1').Select() | Out-Null

if (Test-Path $Salida) { Remove-Item $Salida -Force }
$wb.SaveAs($Salida, 51)
"Dashboard guardado en $Salida"





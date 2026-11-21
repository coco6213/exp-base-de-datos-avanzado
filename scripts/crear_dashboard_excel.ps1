# =====================================================================
# Crea el dashboard BI en Excel conectado al Data Warehouse (DW_Andina).
# Hoja Dashboard:
#   - MEDIDAS (cambian con los filtros): ventas, unidades, precio promedio
#   - KPI 2025 vs META con semáforo: crecimiento, cumplimiento de la meta de ventas,
#     precio promedio y resumen de KPI cumplidos
#   - 4 gráficos y 4 segmentadores (filtros OLAP)
# Hojas de apoyo: Tablas_OLAP (tablas dinámicas) y Datos (conexión ODBC).
# Uso:  powershell -File crear_dashboard_excel.ps1
# =====================================================================
param([string]$Salida = "$PSScriptRoot\..\dashboard\Dashboard_Andina.xlsx")
$ErrorActionPreference = 'Stop'
$Salida = [IO.Path]::GetFullPath($Salida)
New-Item -ItemType Directory -Force (Split-Path $Salida) | Out-Null

# Metas del negocio (definidas en la sección 3.4.1 del informe)
$META_CRECIMIENTO = 0.05     # crecer al menos 5 % frente al año anterior
$META_PRECIO      = 24       # precio promedio por unidad de al menos S/ 24,00
$ANIO = 2025; $ANIO_ANT = 2024

$exe = 'C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE'
if (-not (Get-Process EXCEL -ErrorAction SilentlyContinue)) { Start-Process $exe -ArgumentList '/automation','-Embedding'; Start-Sleep 6 }
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Add()
while ($wb.Worksheets.Count -lt 3) { $wb.Worksheets.Add() | Out-Null }
$dash = $wb.Worksheets.Item(1); $dash.Name = 'Dashboard'
$olap = $wb.Worksheets.Item(2); $olap.Name = 'Tablas_OLAP'
$dat  = $wb.Worksheets.Item(3); $dat.Name  = 'Datos'

# 1) Conexión ODBC al Data Warehouse (Datos > Actualizar todo)
$con = 'ODBC;DRIVER={ODBC Driver 18 for SQL Server};SERVER=localhost;DATABASE=DW_Andina;Trusted_Connection=Yes;TrustServerCertificate=Yes;'
$lo = $dat.ListObjects.Add(0, $con, $true, 1, $dat.Range('A1'))
$lo.QueryTable.CommandText = 'SELECT Anio, Mes, Periodo, Trimestre, Temporada, Region, Sede, Categoria, Producto, Segmento, Unidades, Importe, Ventas FROM dbo.vw_Dashboard'
$lo.QueryTable.Refresh($false) | Out-Null
$lo.Name = 'tblVentas'
$lo.ListColumns.Item('Importe').DataBodyRange.NumberFormat = '#,##0.00'

# 2) Tablas dinámicas (cubo OLAP)
$pc = $wb.PivotCaches().Create(1, 'tblVentas')
function NuevaPivot($nombre, $celda) { $pc.CreatePivotTable($olap.Range($celda), $nombre) }
$xlRow = 1; $xlCol = 2; $xlSum = -4157
$p0 = NuevaPivot 'ptMedidas' 'A3'                               # medidas (responden a todos los filtros)
$f = $p0.AddDataField($p0.PivotFields('Importe'), 'Ventas S/', $xlSum); $f.NumberFormat = '#,##0'
$f = $p0.AddDataField($p0.PivotFields('Unidades'), 'Unidades vendidas', $xlSum); $f.NumberFormat = '#,##0'
$p1 = NuevaPivot 'ptTendencia' 'E3'
$p1.PivotFields('Periodo').Orientation = $xlRow
$f = $p1.AddDataField($p1.PivotFields('Importe'), 'Ventas mensuales S/', $xlSum); $f.NumberFormat = '#,##0'
$p2 = NuevaPivot 'ptRegionAnio' 'I3'
$p2.PivotFields('Region').Orientation = $xlRow
$p2.PivotFields('Anio').Orientation = $xlCol
$f = $p2.AddDataField($p2.PivotFields('Importe'), 'Ventas region', $xlSum); $f.NumberFormat = '#,##0'
$p3 = NuevaPivot 'ptTopProductos' 'P3'
$pf = $p3.PivotFields('Producto'); $pf.Orientation = $xlRow
$f = $p3.AddDataField($p3.PivotFields('Importe'), 'Ventas producto', $xlSum); $f.NumberFormat = '#,##0'
$pf.PivotFilters.Add2(1, $p3.DataFields(1), 10) | Out-Null
$pf.AutoSort(1, $p3.DataFields(1).Name)
$p4 = NuevaPivot 'ptCategoria' 'T3'
$pf = $p4.PivotFields('Categoria'); $pf.Orientation = $xlRow
$f = $p4.AddDataField($p4.PivotFields('Importe'), 'Ventas categoria', $xlSum); $f.NumberFormat = '#,##0'
$pf.AutoSort(2, $p4.DataFields(1).Name)
$pK = NuevaPivot 'ptKPIAnio' 'X3'                               # KPI: ventas por año (no se filtra por año)
$pK.PivotFields('Anio').Orientation = $xlRow
$f = $pK.AddDataField($pK.PivotFields('Importe'), 'Ventas anio', $xlSum); $f.NumberFormat = '#,##0'
$f = $pK.AddDataField($pK.PivotFields('Unidades'), 'Unidades anio', $xlSum); $f.NumberFormat = '#,##0'
$olap.Range('A1').Value2 = 'Tablas OLAP que alimentan el dashboard - fuente: DW_Andina.dbo.vw_Dashboard'
$olap.Range('A1').Font.Bold = $true

# 3) Hoja Dashboard
$dash.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
$dash.Cells.Interior.Color = 0xF7F3F0
$dash.Columns('A').ColumnWidth = 2; $dash.Columns('B:D').ColumnWidth = 8; $dash.Columns('E:U').ColumnWidth = 9
$dash.Rows(1).RowHeight = 32
$t = $dash.Range('B1'); $t.Value2 = 'Distribuidora Andina Comercial S.A.C. - Dashboard de Ventas 2023-2025'
$t.Font.Size = 20; $t.Font.Bold = $true; $t.Font.Color = 0x5A3A1F
$dash.Range('B2').Value2 = 'Fuente: Data Warehouse DW_Andina. Medidas: cambian con los filtros. KPI: año 2025 comparado con su meta. Grupo: Vasquez, Oblea, Gutiérrez, Dominguez.'
$dash.Range('B2').Font.Color = 0x666666

function Etiqueta($celda, $texto, $color) { $c = $dash.Range($celda); $c.Value2 = $texto; $c.Font.Bold = $true; $c.Font.Size = 11; $c.Font.Color = $color }
function Celda($rango, $valor, $size, $color, [switch]$bold, [switch]$formula) {
  $r = $dash.Range($rango); $null = $r.Merge(); $r.HorizontalAlignment = -4108; $r.VerticalAlignment = -4108; $r.WrapText = $true
  if ($formula) { $r.Cells.Item(1, 1).Formula = $valor } else { $r.Cells.Item(1, 1).Value2 = $valor }
  $r.Font.Size = $size; $r.Font.Color = $color; $r.Font.Bold = [bool]$bold
  return ,$r
}
function Marco($rango, $borde) { $r = $dash.Range($rango); $r.Interior.Color = 0xFFFFFF; $null = $r.BorderAround(1, 2, [Reflection.Missing]::Value, $borde) }

# --- MEDIDAS (filas 3 a 5) ---
Etiqueta 'F3' 'MEDIDAS (cambian con los filtros)' 0x8B5A1F
$med = @(
  @('F', 'I', 'VENTAS TOTALES (S/)',       '=GETPIVOTDATA("Ventas S/",Tablas_OLAP!$A$3)', '"S/ "#,##0'),
  @('K', 'N', 'UNIDADES VENDIDAS',         '=GETPIVOTDATA("Unidades vendidas",Tablas_OLAP!$A$3)', '#,##0'),
  @('P', 'S', 'PRECIO PROMEDIO POR UNIDAD',  '=GETPIVOTDATA("Ventas S/",Tablas_OLAP!$A$3)/GETPIVOTDATA("Unidades vendidas",Tablas_OLAP!$A$3)', '"S/ "0.00'))
foreach ($m in $med) {
  Marco "$($m[0])4:$($m[1])5" 0xD9D9D9
  [void](Celda "$($m[0])4:$($m[1])4" $m[2] 9 0x555555 -bold)
  $v = Celda "$($m[0])5:$($m[1])5" $m[3] 18 0x8B5A1F -bold -formula; $v.NumberFormat = $m[4]
}
$dash.Rows(4).RowHeight = 18; $dash.Rows(5).RowHeight = 30

# --- KPI 2025 vs META (filas 7 a 10) ---
Etiqueta 'F7' "KPI $ANIO vs META (semáforo: verde = cumple, rojo = no cumple)" 0x1F7A1F
$V  = "GETPIVOTDATA(""Ventas anio"",Tablas_OLAP!`$X`$3,""Anio"",$ANIO)"
$V0 = "GETPIVOTDATA(""Ventas anio"",Tablas_OLAP!`$X`$3,""Anio"",$ANIO_ANT)"
$U  = "GETPIVOTDATA(""Unidades anio"",Tablas_OLAP!`$X`$3,""Anio"",$ANIO)"
$kpi = @(
  @('F', 'I', "KPI 1 · Crecimiento de ventas $ANIO vs $ANIO_ANT", "=$V/$V0-1", '+0.00%;-0.00%', "=""Meta: ≥ +5,00 %""", ('F9>=' + $META_CRECIMIENTO)),
  @('K', 'N', "KPI 2 · Cumplimiento de la meta de ventas $ANIO", "=$V/($V0*(1+$META_CRECIMIENTO))", '0.0%', "=""Meta: 100 % = S/ ""&TEXT($V0*(1+$META_CRECIMIENTO),""#,##0"")", 'K9>=1'),
  @('P', 'S', "KPI 3 · Precio promedio por unidad $ANIO", "=$V/$U", '"S/ "0.00', "=""Meta: ≥ S/ 24,00""", ('P9>=' + $META_PRECIO)))
foreach ($k in $kpi) {
  Marco "$($k[0])8:$($k[1])10" 0xBFBFBF
  [void](Celda "$($k[0])8:$($k[1])8" $k[2] 10 0x555555 -bold)
  $v = Celda "$($k[0])9:$($k[1])9" $k[3] 20 0x333333 -bold -formula; $v.NumberFormat = $k[4]
  [void](Celda "$($k[0])10:$($k[1])10" $k[5] 9 0x555555 -formula)
  $cell = $dash.Range("$($k[0])9")
  $umbral = [string]$k[6].Split('=')[1]; $ok = $cell.FormatConditions.Add(1, 7, $umbral); $ok.Font.Color = 0x1F7A1F; $ok.Interior.Color = 0xE6F4E6
  $no = $cell.FormatConditions.Add(1, 6, $umbral); $no.Font.Color = 0x2B39C0; $no.Interior.Color = 0xE6E6FA
}
# Estado (✔ / ✖) debajo de la meta, en la fila 10 junto a la meta
foreach ($k in $kpi) {
  $c = $dash.Range("$($k[0])10"); $c.Formula = "=$($c.Formula.Substring(1))&""   ""&IF($($k[6]),""✔ CUMPLE"",""✖ NO CUMPLE"")"
}
# KPI 4: resumen
Marco 'U8:X10' 0xBFBFBF
[void](Celda 'U8:X8' 'Resumen de KPI' 10 0x555555 -bold)
$r = Celda 'U9:X9' "=((F9>=$META_CRECIMIENTO)+(K9>=1)+(P9>=$META_PRECIO))&"" de 3""" 20 0x333333 -bold -formula
[void](Celda 'U10:X10' "KPI cumplidos en $ANIO" 9 0x555555)
$dash.Rows(8).RowHeight = 30; $dash.Rows(9).RowHeight = 30; $dash.Rows(10).RowHeight = 28

# 4) Gráficos (desde la fila 12)
function Grafico($pt, $tipo, $titulo, $left, $top, $w, $h) {
  $sh = $dash.Shapes.AddChart2(-1, $tipo, $left, $top, $w, $h)
  $sh.Chart.SetSourceData($pt.TableRange1)
  $sh.Chart.HasTitle = $true; $sh.Chart.ChartTitle.Text = $titulo
  $sh.Chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = 12
  try { $sh.Chart.ShowAllFieldButtons = $false } catch {}
  return $sh.Chart
}
$L = $dash.Range('F12').Left; $T = $dash.Range('F12').Top
$c1 = Grafico $p1 4 'Tendencia mensual de ventas (S/) - picos en campaña Nov-Dic' $L $T 520 230
$c1.HasLegend = $false
$c2 = Grafico $p2 52 'Ventas por región y año (S/) - análisis OLAP' ($L + 530) $T 420 230
$c2.ApplyDataLabels(2); for ($i = 1; $i -le $c2.SeriesCollection().Count; $i++) { $dl = $c2.SeriesCollection($i).DataLabels(); $dl.Font.Color = 0xFFFFFF; $dl.Font.Size = 8; $dl.Font.Bold = $true }; $c2.HasLegend = $true; $c2.Legend.Position = -4107
$c3 = Grafico $p3 57 'Top 10 productos por ventas (S/)' $L ($T + 240) 520 240
$c3.HasLegend = $false; $c3.Axes(2).MinimumScale = 0; $c3.ApplyDataLabels(2)
$c4 = Grafico $p4 5 'Participación por categoría' ($L + 530) ($T + 240) 420 240
$c4.ApplyDataLabels(3); $c4.HasLegend = $true; $c4.Legend.Position = -4152

# 5) Segmentadores: medidas y gráficos responden a los 4 filtros; los KPI a región, segmento y temporada
$top = $dash.Range('B4').Top
foreach ($campo in 'Anio', 'Region', 'Segmento', 'Temporada') {
  $sc = $wb.SlicerCaches.Add2($p0, $campo)
  foreach ($p in $p1, $p2, $p3, $p4) { $sc.PivotTables.AddPivotTable($p) }
  if ($campo -ne 'Anio') { $sc.PivotTables.AddPivotTable($pK) }
  $alto = @{Anio = 110; Region = 110; Segmento = 110; Temporada = 90 }[$campo]
  $sc.Slicers.Add($dash, [Type]::Missing, "sl$campo", $campo, $top, 6, 150, $alto) | Out-Null
  $top += $alto + 8
}
$dash.Range('A1').Select() | Out-Null
if (Test-Path $Salida) { Remove-Item $Salida -Force }
$wb.SaveAs($Salida, 51)
"Dashboard guardado en $Salida"

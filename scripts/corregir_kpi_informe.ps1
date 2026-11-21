# Corrige en el informe Word la parte de KPI: distingue MEDIDAS (sin meta) de KPI (con meta y semáforo),
# actualiza la tabla 3.4.1, los textos, las figuras 7 y 8 y la prueba T-09.
param([string]$Doc)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$d = [Runtime.InteropServices.Marshal]::BindToMoniker($Doc)
"Word: " + $d.Application.Path

function Reemplazar($viejo, $nuevo) {
  $rng = $d.Content; $f = $rng.Find; $f.ClearFormatting(); $f.Text = $viejo; $f.MatchCase = $true; $f.Forward = $true; $f.Wrap = 0
  if ($f.Execute()) { $rng.Text = $nuevo; "OK: $($viejo.Substring(0, [Math]::Min(50, $viejo.Length)))" } else { "NO ENCONTRADO: $($viejo.Substring(0, [Math]::Min(50, $viejo.Length)))" }
}
function Parrafo($inicio) {
  $rng = $d.Content; $f = $rng.Find; $f.ClearFormatting(); $f.Text = $inicio; $f.MatchCase = $true; $f.Forward = $true; $f.Wrap = 0
  if ($f.Execute()) { return $rng.Paragraphs.Item(1) } else { return $null }
}
function TextoParrafo($inicio, $nuevo) {
  $p = Parrafo $inicio
  if ($p -eq $null) { "NO ENCONTRADO: $inicio"; return }
  $d.Range($p.Range.Start, $p.Range.End - 1).Text = $nuevo; "OK párrafo: $inicio"
}

# 1) Resumen y abstract
Reemplazar 'y un dashboard en Excel con 4 KPI, 4 gráficos y filtros OLAP' 'y un dashboard en Excel con medidas, 3 KPI con meta y semáforo, 4 gráficos y filtros OLAP'
Reemplazar 'and an Excel dashboard with 4 KPIs, 4 charts and OLAP filters' 'and an Excel dashboard with measures, 3 KPIs with targets and traffic lights, 4 charts and OLAP filters'

# 2) Stakeholders (1.5), fundamentos (2.2 y 2.4), requisitos (3.1) y alternativas (3.2)
Reemplazar 'KPI ventas totales y crecimiento % disponibles a diario; meta ≥ 4 KPI en el dashboard' 'KPI de crecimiento y de cumplimiento de la meta de ventas visibles con semáforo (3 KPI con meta)'
TextoParrafo 'Indicadores (KPI): números clave' 'Indicadores (KPI): un KPI compara un resultado con su meta y dice si se cumplió (por ejemplo, crecimiento anual frente a la meta de +5 %). Un número sin meta, como las ventas totales, es una medida: describe la situación, pero no dice si es buena o mala.'
Reemplazar 'Excel es suficiente para 4 KPI y filtros OLAP' 'Excel es suficiente para 3 KPI con meta, semáforo y filtros OLAP'
Reemplazar '≥ 4 visualizaciones; filtros por año, región, segmento y temporada' '≥ 3 KPI con meta y semáforo; ≥ 4 visualizaciones; filtros por año, región, segmento y temporada'
Reemplazar 'Seleccionada: cumple 4 KPI + OLAP sin costo' 'Seleccionada: cumple KPI con meta + OLAP sin costo'

# 3) Tabla de indicadores (3.4.1): medidas vs KPI
$tk = $null
foreach ($t in $d.Tables) { try { if (($t.Cell(1, 3).Range.Text -replace '[\r\a]', '') -like 'KPI Y F*') { $tk = $t; break } } catch {} }
if ($tk -ne $null) {
  $filas = @(
    @('PREGUNTA DE NEGOCIO', 'USUARIO / DECISIÓN', 'TIPO, NOMBRE Y FÓRMULA', 'GRANO / FRECUENCIA', 'FUENTE', 'META Y RESULTADO 2025'),
    @('¿Crecemos o no?', 'Gerencia / estrategia', 'KPI 1 · Crecimiento % = Ventas 2025 / Ventas 2024 − 1', 'Año, región / mensual', 'FactVentas + DimTiempo', 'Meta ≥ +5 % · Real −0,41 % · ✖ No cumple'),
    @('¿Cumplimos la meta de ventas?', 'Gerencia / metas comerciales', 'KPI 2 · Cumplimiento = Ventas 2025 / (Ventas 2024 × 1,05)', 'Año, región / mensual', 'FactVentas + DimTiempo', 'Meta 100 % (S/ 63,13 M) · Real 94,8 % · ✖ No cumple'),
    @('¿A qué precio vendemos?', 'Comercial / precios', 'KPI 3 · Precio promedio = SUM(Importe) / SUM(Cantidad)', 'Año, categoría / semanal', 'FactVentas', 'Meta ≥ S/ 24,00 · Real S/ 24,77 · ✔ Cumple'),
    @('¿Cuánto vendemos?', 'Gerencia general / contexto', 'Medida · Ventas totales = SUM(Importe)', 'Mes / diaria', 'FactVentas', 'Sin meta (dato de contexto): S/ 180,1 M en 2023-2025'),
    @('¿Cuánto volumen movemos?', 'Logística / compras', 'Medida · Unidades = SUM(Cantidad)', 'Mes, producto / diaria', 'FactVentas', 'Sin meta: 7,28 M de unidades'),
    @('¿Qué productos impulsar?', 'Comercial / promociones', 'Medida · Ranking top 10 por importe', 'Producto / semanal', 'FactVentas + DimProducto', 'Sin meta: apoya las promociones'),
    @('¿Cuándo hay más demanda?', 'Logística / campaña', 'Medida · Ventas por mes en campaña vs regular', 'Mes / anual', 'DimTiempo.EsCampana', 'Sin meta: campaña = 2,5 veces más'))
  while ($tk.Rows.Count -lt $filas.Count) { [void]$tk.Rows.Add() }
  while ($tk.Rows.Count -gt $filas.Count) { $tk.Rows.Item($tk.Rows.Count).Delete() }
  for ($i = 0; $i -lt $filas.Count; $i++) { for ($j = 0; $j -lt 6; $j++) { $tk.Cell($i + 1, $j + 1).Range.Text = $filas[$i][$j] } }
  for ($i = 2; $i -le 4; $i++) { $tk.Rows.Item($i).Shading.BackgroundPatternColor = 15132390 }   # filas KPI resaltadas
  # texto explicativo antes de la tabla
  $h = Parrafo '3.4.1 Requisitos, usuarios e indicadores'
  $pos = $h.Range.End - 1
  $txt = 'Se distinguen dos tipos de indicadores. Las medidas describen la situación (por ejemplo, cuánto se vendió) y no tienen meta. Los KPI comparan el resultado del año 2025 con una meta del negocio y muestran con un semáforo si se cumplió (verde) o no (rojo). Las filas resaltadas son los 3 KPI.'
  $d.Range($pos, $pos).InsertAfter("`r" + $txt)
  $np = $d.Range($pos + 1, $pos + 1).Paragraphs.Item(1)
  $np.Style = $d.Styles.Item('Normal'); $np.Range.Font.Reset(); $np.Alignment = 3
  "OK tabla 3.4.1 (medidas vs KPI)"
} else { "NO ENCONTRADA tabla KPI" }

# 4) Sección 3.4.4: contenido, interpretación, figuras
TextoParrafo 'Contenido: 4 KPI' 'Contenido: 3 medidas que cambian con los filtros (ventas totales, unidades y precio promedio), 3 KPI del año 2025 comparados con su meta y con semáforo (crecimiento de ventas, cumplimiento de la meta de ventas y precio promedio), un resumen de KPI cumplidos, 4 gráficos (tendencia mensual, región por año, top 10 productos y participación por categoría) y 4 segmentadores OLAP (año, región, segmento y temporada).'
TextoParrafo '• Medidas Dinamicas:' '• Medidas: en 3 años (2023-2025) se vendieron S/ 180,1 millones y 7,28 millones de unidades (precio promedio S/ 24,75). Describen la situación y sirven de base para fijar metas.'
$p = Parrafo '• Medidas: en 3 años'
if ($p -ne $null) {
  $pos = $p.Range.End - 1
  $kpiTxt = '• KPI 2025 vs meta: el crecimiento fue −0,41 % (meta ≥ +5 %, no cumple), el cumplimiento de la meta de ventas fue 94,8 % (meta 100 % = S/ 63,13 millones, no cumple) y el precio promedio fue S/ 24,77 (meta ≥ S/ 24,00, cumple). Resultado: 1 de 3 KPI cumplidos. Decisión: la gerencia debe reforzar las ventas con promociones y nuevos clientes; el precio se mantiene.'
  $d.Range($pos, $pos).InsertAfter("`r" + $kpiTxt); "OK viñeta KPI"
}
TextoParrafo 'Figura 7.' 'Figura 7. Dashboard BI de ventas 2023-2025: filtros (1), medidas (2), KPI 2025 vs meta con semáforo (3), tendencia (4) y análisis OLAP región x año (5) (evidencia E-07).'
TextoParrafo 'Figura 8.' 'Figura 8. Operación OLAP "dice": filtro año 2024 y región Trujillo; las medidas se recalculan (S/ 11,96 millones) y los KPI muestran a Trujillo frente a su meta (evidencia E-07b).'

# 5) Imágenes de las figuras 7 y 8 (y anotaciones flotantes antiguas)
function CambiarImagen($caption, $archivo) {
  $c = Parrafo $caption
  $prev = $c.Previous(1)
  if ($prev.Range.InlineShapes.Count -lt 1) { "NO hay imagen antes de: $caption"; return }
  $old = $prev.Range.InlineShapes.Item(1); $w = $old.Width
  $r = $old.Range; $old.Delete()
  $pic = $d.InlineShapes.AddPicture($archivo, $false, $true, $r)
  $ratio = $pic.Height / $pic.Width; $pic.Width = $w; $pic.Height = $w * $ratio
  "OK imagen: $caption"
}
# borrar anotaciones flotantes antiguas entre "Contenido:" y la figura 8
$ini = (Parrafo 'Contenido: 3 medidas').Range.Start; $fin = (Parrafo 'Figura 8.').Range.End
$borrar = @(); foreach ($s in $d.Shapes) { $a = $s.Anchor.Start; if ($a -ge $ini -and $a -le $fin) { $borrar += $s } }
foreach ($s in $borrar) { $s.Delete() }; "Anotaciones flotantes borradas: $($borrar.Count)"
CambiarImagen 'Figura 7.' (Join-Path $root 'evidencias\E-07_dashboard_general.png')
CambiarImagen 'Figura 8.' (Join-Path $root 'evidencias\E-07b_dashboard_filtro_trujillo_2024.png')

# 6) Prueba T-09 (sección 3.7)
Reemplazar 'Filtro 2025 + Lima en el dashboard' 'Filtro 2024 + Trujillo en el dashboard'
Reemplazar 'KPI y gráficos se recalculan' 'Medidas, KPI y gráficos se recalculan'
Reemplazar 'S/ 29 973 672 – PASA' 'S/ 11 956 809 – PASA'
Reemplazar 'dashboard con 4 visualizaciones y filtros funcionando' 'dashboard con 3 KPI con meta, 4 visualizaciones y filtros funcionando'

if ($d.TablesOfContents.Count -gt 0) { $d.TablesOfContents.Item(1).Update() }
$d.Save()
"Páginas: " + $d.ComputeStatistics(2)

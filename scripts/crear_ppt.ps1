# Crea la presentación del avance (semana 6) con Microsoft PowerPoint. 14 diapositivas + notas del orador.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$pres = $pp.Presentations.Add($false)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
function RGBc($r, $g, $b) { return [int]($r + 256 * $g + 65536 * $b) }
$NAVY = RGBc 31 58 90; $BLUE = RGBc 31 90 139; $GOLD = RGBc 184 134 11; $GRAY = RGBc 90 90 90; $LIGHT = RGBc 242 246 250; $RED = RGBc 192 57 43; $GREEN = RGBc 31 122 58
$script:n = 0

function Txt($s, $l, $t, $w, $h, $text, $size, $color, [switch]$bold, [int]$align = 1) {
  $b = $s.Shapes.AddTextbox(1, $l, $t, $w, $h); $tf = $b.TextFrame; $tf.WordWrap = -1
  $r = $tf.TextRange; $r.Text = $text; $r.Font.Name = 'Segoe UI'; $r.Font.Size = $size; $r.Font.Color.RGB = $color; $r.Font.Bold = [int][bool]$bold
  $r.ParagraphFormat.Alignment = $align
  return $b
}
function Slide($title, $notes) {
  $script:n++
  $s = $pres.Slides.Add($script:n, 12)
  $bar = $s.Shapes.AddShape(1, 0, 0, 12, 540); $bar.Fill.ForeColor.RGB = $BLUE; $bar.Line.Visible = 0
  [void](Txt $s 36 22 880 50 $title 28 $NAVY -bold)
  $ln = $s.Shapes.AddLine(36, 74, 924, 74); $ln.Line.ForeColor.RGB = $GOLD; $ln.Line.Weight = 2
  [void](Txt $s 36 510 700 22 'DataShield 360 · Distribuidora Andina · Grupo: Channel, Yin, Gutiérrez, Pirsin' 10 $GRAY)
  [void](Txt $s 860 510 64 22 "$script:n" 10 $GRAY -align 3)
  if ($notes) { $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes }
  return $s
}
function Bullets($s, $l, $t, $w, $h, $lines, $size = 18) {
  $b = Txt $s $l $t $w $h ($lines -join "`r") $size (RGBc 40 40 40)
  $r = $b.TextFrame.TextRange; $r.ParagraphFormat.Bullet.Visible = -1; $r.ParagraphFormat.Bullet.Character = 8226
  $r.ParagraphFormat.SpaceAfter = 6
  return $b
}
function Img($s, $path, $l, $t, $maxW, $maxH) {
  $p = $s.Shapes.AddPicture((Join-Path $root $path), 0, -1, $l, $t)
  $k = [Math]::Min($maxW / $p.Width, $maxH / $p.Height); $p.LockAspectRatio = -1; $p.Width = $p.Width * $k
  $p.Left = $l + ($maxW - $p.Width) / 2; $p.Top = $t + ($maxH - $p.Height) / 2
  $p.Line.Visible = -1; $p.Line.ForeColor.RGB = (RGBc 200 200 200)
  return $p
}
function Stat($s, $l, $t, $w, $num, $label, $color) {
  $box = $s.Shapes.AddShape(5, $l, $t, $w, 120); $box.Fill.ForeColor.RGB = $LIGHT; $box.Line.ForeColor.RGB = (RGBc 210 220 230)
  [void](Txt $s $l ($t + 12) $w 50 $num 34 $color -bold -align 2)
  [void](Txt $s ($l + 8) ($t + 64) ($w - 16) 54 $label 13 $GRAY -align 2)
}

# 1. Portada
$script:n++
$s = $pres.Slides.Add(1, 12)
$bg = $s.Shapes.AddShape(1, 0, 0, 960, 540); $bg.Fill.ForeColor.RGB = $NAVY; $bg.Line.Visible = 0
$st = $s.Shapes.AddShape(1, 60, 300, 120, 4); $st.Fill.ForeColor.RGB = $GOLD; $st.Line.Visible = 0
[void](Txt $s 60 90 840 40 'BASE DE DATOS AVANZADAS Y BIG DATA · CIIN1487R' 16 (RGBc 200 215 230) -bold)
[void](Txt $s 60 140 840 90 'DataShield 360: ecosistema de datos seguro e inteligente' 36 (RGBc 255 255 255) -bold)
[void](Txt $s 60 240 840 40 'Avance semana 6 · Inteligencia de Negocios y Visualización de Datos' 20 (RGBc 230 200 120))
[void](Txt $s 60 320 840 120 ("Caso: Distribuidora Andina Comercial S.A.C. (Lima, Arequipa y Trujillo)`rGrupo: Channel · Yin · Gutiérrez · Pirsin`rDocente: Ramirez Visa Adan Elvis · NRC 3814 · 2026-2") 16 (RGBc 255 255 255))
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Presentamos el avance hasta Inteligencia de Negocios y Visualización. Hoy mostramos la base de datos, la seguridad, el Data Warehouse, el dashboard y una prueba de rendimiento. Big Data y MongoDB vienen en la entrega final.'

# 2. Problema
$s = Slide 'El problema de la empresa' 'La empresa decide con Excel hecho a mano una vez al mes. No sabe quién cambia los datos, no tiene respaldos formales y la autoridad le pide cumplir la ley de datos personales en 60 días.'
[void](Bullets $s 40 95 520 380 @('Reportes de Excel hechos a mano cada fin de mes: decisiones con 1 mes de retraso', 'La base SQL Server no tiene procedimientos, auditoría ni respaldos formales', 'Datos con errores: ventas anuladas mezcladas, nombres sucios, emails vacíos', 'La autoridad exige cumplir la Ley N.° 29733 en 60 días') 18)
Stat $s 600 100 150 '78 %' 'empresas con pérdida o corrupción de datos (caso)' $RED
Stat $s 770 100 150 '91 %' 'medianas empresas deciden sin análisis formal' $RED
Stat $s 600 250 320 '60 días' 'plazo para demostrar cumplimiento normativo' $GOLD

# 3. Alcance
$s = Slide 'Qué incluye este avance (y qué falta)' 'Este avance cubre hasta BI y visualización, como pidió el profesor. Lo de la derecha es lo que entregaremos al final del curso.'
[void](Txt $s 40 92 430 30 'Incluido hoy (semana 6)' 20 $GREEN -bold)
[void](Bullets $s 40 128 430 340 @('Base transaccional de 11 tablas con datos simulados (200 000 ventas)', '3 procedimientos y 2 triggers con COMMIT / ROLLBACK', '3 roles, auditoría y respaldos completo, diferencial y log', 'ETL con 5 reglas de limpieza y Data Warehouse (esquema estrella)', 'Dashboard BI en Excel con filtros OLAP', 'Prueba de viabilidad del rendimiento (3 escenarios)') 16)
[void](Txt $s 500 92 430 30 'Entrega final (semanas 7-16)' 20 $GOLD -bold)
[void](Bullets $s 500 128 420 340 @('MongoDB: comentarios y rutas en JSON (CRUD)', 'PySpark: ranking y estacionalidad con ≥ 10 000 registros', 'Integrar resultados de Spark al dashboard', 'Experimento completo, resultados y conclusiones (cap. IV-VIII)', 'Informe de Medición completo') 16)

# 4. Arquitectura
$s = Slide 'Arquitectura de la solución' 'Así viajan los datos: se registran en la base del día a día, el ETL los limpia y los pasa al Data Warehouse, y el dashboard los muestra a la gerencia. Abajo, punteado, lo que agregaremos con Big Data.'
[void](Img $s 'diagramas\E-05c_arquitectura.png' 40 85 880 415)

# 5. Base de datos y automatización
$s = Slide 'Base de datos y automatización' 'Un procedimiento es una receta guardada que la base ejecuta siempre igual. Un trigger es un vigilante que anota cada cambio. Si una venta no tiene stock, la transacción deshace todo: no se guarda nada a medias.'
[void](Img $s 'diagramas\E-05a_diagrama_er.png' 30 85 560 410)
[void](Bullets $s 610 95 320 400 @('11 tablas con llaves y reglas CHECK (ej.: stock nunca negativo)', 'usp_RegistrarVenta: todo o nada (COMMIT / ROLLBACK)', 'usp_ActualizarInventario y usp_GenerarAlertasStock', '2 triggers anotan quién cambió qué y cuándo (Cliente y Producto)', 'Prueba: venta sin stock → ROLLBACK, 0 ventas a medias') 15)

# 6. Seguridad
$s = Slide 'Seguridad y respaldos' 'Cada persona solo puede hacer lo que su trabajo necesita. El analista no puede ver email, teléfono ni RUC. Los respaldos permiten perder como máximo 15 minutos de información.'
[void](Img $s 'evidencias\E-06c_roles_ssms.png' 30 90 520 300)
[void](Bullets $s 570 95 360 300 @('Administrador: control total', 'Analista: solo lectura y SIN datos personales', 'Operativo: solo usa los procedimientos, no puede borrar', 'Backups: completo (domingo), diferencial (diario), log (15 min)', 'Ley 29733 e ISO/IEC 27001: matriz de cumplimiento') 15)
[void](Txt $s 30 400 520 60 'En rojo: el sistema le niega al analista el acceso al email del cliente.' 13 $GRAY)

# 7. DW y ETL
$s = Slide 'Data Warehouse y ETL' 'El Data Warehouse es una copia ordenada solo para analizar. El ETL extrae, limpia y carga. Comprobamos que el total vendido sea igual al céntimo en los dos lados.'
[void](Img $s 'diagramas\E-05b_esquema_estrella.png' 30 85 480 360)
[void](Bullets $s 530 95 400 260 @('Esquema estrella: 1 tabla de hechos + 4 dimensiones (cuándo, qué, quién, dónde)', '5 reglas de limpieza: nombres, anuladas, fechas, datos personales, valores', '599 808 líneas → 17 667 rechazadas → 582 141 cargadas', 'Índice columnar: 6 MB frente a 35 MB') 15)
Stat $s 560 370 340 'S/ 0,00' 'diferencia entre la base diaria y el Data Warehouse (S/ 180 113 494,64)' $GREEN

# 8. Dashboard
$s = Slide 'Dashboard BI (Excel conectado al Data Warehouse)' 'Este es el dashboard. Arriba los KPI, a la izquierda los filtros. Al hacer clic en un filtro, todo se recalcula: eso es el análisis OLAP.'
[void](Img $s 'evidencias\E-07_dashboard.png' 40 85 880 415)

# 9. Decisiones
$s = Slide 'Qué nos dice el dashboard' 'Lo más importante para la gerencia: noviembre y diciembre venden 2,5 veces más, y las ventas no crecen. Con el filtro vemos cualquier región o año en un clic.'
[void](Img $s 'evidencias\E-07b_dashboard_filtro_lima_2025.png' 30 90 500 300)
[void](Bullets $s 550 95 380 380 @('Nov-Dic venden ≈ S/ 10 millones al mes vs ≈ S/ 4 millones (2,5 veces) → preparar stock desde octubre', 'Lima 50 %, Arequipa 30 %, Trujillo 20 %', 'Ventas planas: +0,03 % (2024) y −0,41 % (2025) → se necesitan nuevas acciones comerciales', 'Top 10 productos para priorizar stock y promociones') 15)
[void](Txt $s 30 400 500 40 'Filtro OLAP: año 2025 + Lima → S/ 29,97 millones' 13 $GRAY)

# 10. Experimento
$s = Slide 'Prueba de viabilidad: ¿el Data Warehouse acelera los reportes?' 'Medimos 4 preguntas de la gerencia en 3 escenarios, 10 veces cada una. Las barras azules son el Data Warehouse: siempre las más bajas, es decir, las más rápidas.'
[void](Img $s 'evidencias\E-09b_grafico_tiempos.png' 30 85 600 320)
[void](Bullets $s 650 95 280 320 @('E1: base actual', 'E2: base + 4 índices', 'E3: Data Warehouse columnar', '4 consultas × 10 repeticiones + 1 en frío', 'Mismo resultado en los 3 (checksum)') 15)
[void](Txt $s 30 415 900 60 'Hipótesis: E3 al menos 5 veces más rápido que E1 en los reportes → SE CUMPLE (9,5 a 11,5 veces).' 16 $NAVY -bold)

# 11. Resultados
$s = Slide 'Resultados clave' 'Cuatro números para recordar: el Data Warehouse es unas 10 veces más rápido, lee 95 % menos, el ETL cuadra al céntimo y todas las pruebas pasaron.'
Stat $s 40 110 210 '10x' 'reportes más rápidos con el Data Warehouse (9,5 a 11,5 veces)' $BLUE
Stat $s 265 110 210 '−95 %' 'páginas leídas en los reportes' $BLUE
Stat $s 490 110 210 'S/ 0,00' 'diferencia OLTP vs DW' $GREEN
Stat $s 715 110 210 '10 / 10' 'pruebas funcionales y de seguridad PASA' $GREEN
[void](Bullets $s 40 270 880 200 @('Los índices tradicionales solo aceleran la búsqueda de un cliente (105 veces), no los reportes', 'En la primera consulta (memoria vacía) el Data Warehouse tarda más: luego es el más rápido', 'Decisión: el dashboard se conecta al Data Warehouse, no a la base del día a día') 16)

# 12. Mejoras y riesgos
$s = Slide 'Mejoras aplicadas y riesgos' 'Encontramos errores y los corregimos: un sesgo en los datos simulados y un problema de reproducibilidad. También identificamos los riesgos para la entrega final.'
[void](Txt $s 40 92 430 30 'Ciclo de mejora (retroalimentación)' 18 $GREEN -bold)
[void](Bullets $s 40 126 430 340 @('Docente: el avance es el Informe de Proyecto Final hasta BI → se reorganizó', 'Sesgo: en Lima todos eran "bodega" → generador v2 + validación V13', 'IDs no reproducibles → ORDER BY + huella digital del dataset', 'Prueba T-08 en mal orden → script 11 separado') 15)
[void](Txt $s 500 92 430 30 'Riesgos principales' 18 $RED -bold)
[void](Bullets $s 500 126 420 340 @('Python/PySpark no instalado (alta) → instalar o usar Colab, +3 días', 'MongoDB no instalado (alta) → Community o Atlas, +2 días', 'Variación de tiempos (media) → 10 repeticiones y desviación estándar', 'Datos personales (baja, alto impacto) → roles y DW sin datos personales') 15)

# 13. Próximos pasos
$s = Slide 'Cronograma y próximos pasos' 'La ruta crítica es Big Data con PySpark. Le dimos 3 días extra de contingencia, gracias a que eliminamos la tarea de optimizar índices, que el experimento mostró que no era útil.'
[void](Bullets $s 40 95 880 380 @('H1-H4 (semanas 5-6): análisis, base de datos, seguridad, DW, dashboard y experimento → COMPLETADO', 'H5 (23/11 – 04/12): MongoDB con CRUD y comparativa SQL vs NoSQL (+2 días contingencia)', 'H6 (07/12 – 23/12): PySpark con ≥ 10 000 registros (ruta crítica, +3 días contingencia)', 'H7 (04/01 – 15/01): integrar Spark al dashboard y experimento de escalabilidad', 'H8-H9 (enero): Informe de Medición, informe final y sustentación', 'Pendiente: publicar el repositorio (etiqueta v0.6-avance-s6) y compartirlo con el docente') 17)

# 14. Conclusiones
$s = Slide 'Conclusiones del avance' 'Cerramos con lo esencial: separar la base del día a día del Data Warehouse fue la decisión correcta, y está demostrado con datos.'
[void](Bullets $s 40 95 880 380 @('Separar la base del día a día del Data Warehouse es la decisión correcta: reportes ~10 veces más rápidos', 'Agregar índices no resuelve los reportes: solo ayuda a búsquedas puntuales', 'El ETL entrega información confiable: cuadra al céntimo con el origen', 'El dashboard muestra decisiones claras: stock para campaña y metas comerciales', 'La base ahora es segura: transacciones, auditoría, roles y respaldos probados') 19)

$out = "$root\informe\Grupo_CIIN1487R_Avance_S6.pptx"
if (Test-Path $out) { Remove-Item $out -Force }
$pres.SaveAs($out)
$pres.SaveAs("$root\informe\Grupo_CIIN1487R_Avance_S6_ppt.pdf", 32)
"Diapositivas: $($pres.Slides.Count) -> $out"
$pres.Close()

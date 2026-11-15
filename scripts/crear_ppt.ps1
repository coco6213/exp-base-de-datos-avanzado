# Crea la presentación del avance (semana 6) en base al informe AVANCE-EXP-DEDATOSAVANZADO.
# 15 diapositivas con estructura uniforme: sección + título + contenido + "Idea clave" + notas del orador.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
try { $pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
catch { Start-Process 'C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE' -ArgumentList '/automation', '-Embedding'; Start-Sleep 8; $pp = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') }
$pres = $pp.Presentations.Add($false)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
function C($r, $g, $b) { [int]($r + 256 * $g + 65536 * $b) }
$NAVY = C 31 58 90; $BLUE = C 31 90 139; $GOLD = C 184 134 11; $GRAY = C 100 100 100; $TEXT = C 35 35 35
$LIGHT = C 238 244 250; $GREEN = C 31 122 58; $RED = C 192 57 43; $WHITE = C 255 255 255
$script:n = 0

function Txt($s, $l, $t, $w, $h, $text, $size, $color, [switch]$bold, [int]$align = 1) {
  $b = $s.Shapes.AddTextbox(1, $l, $t, $w, $h); $tf = $b.TextFrame; $tf.WordWrap = -1; $tf.AutoSize = 0
  $tf.MarginLeft = 4; $tf.MarginRight = 4
  $r = $tf.TextRange; $r.Text = $text; $r.Font.Name = 'Segoe UI'; $r.Font.Size = $size; $r.Font.Color.RGB = $color
  $r.Font.Bold = [int][bool]$bold; $r.ParagraphFormat.Alignment = $align
  return $b
}
function Box($s, $l, $t, $w, $h, $fill, $line) {
  $b = $s.Shapes.AddShape(1, $l, $t, $w, $h); $b.Fill.ForeColor.RGB = $fill
  if ($line -ne $null) { $b.Line.ForeColor.RGB = $line; $b.Line.Weight = 1 } else { $b.Line.Visible = 0 }
  return $b
}
function Slide($section, $title, $idea, $notes) {
  $script:n++
  $s = $pres.Slides.Add($script:n, 12)
  [void](Box $s 0 0 960 6 $BLUE $null)
  [void](Txt $s 36 14 700 20 $section.ToUpper() 11 $GOLD -bold)
  [void](Txt $s 36 32 888 42 $title 26 $NAVY -bold)
  $ln = $s.Shapes.AddLine(40, 78, 920, 78); $ln.Line.ForeColor.RGB = $GOLD; $ln.Line.Weight = 1.5
  if ($idea) {
    [void](Box $s 40 446 880 46 $LIGHT $null); [void](Box $s 40 446 6 46 $GOLD $null)
    $t = Txt $s 54 450 860 40 "Idea clave: $idea" 14 $NAVY
    $t.TextFrame.VerticalAnchor = 3; $t.TextFrame.TextRange.Characters(1, 11).Font.Bold = -1
  }
  [void](Txt $s 36 508 760 20 'DataShield 360 · Distribuidora Andina · Avance semana 6 · Channel, Yin, Gutiérrez, Pirsin' 9 $GRAY)
  [void](Txt $s 860 508 64 20 "$script:n / 15" 9 $GRAY -align 3)
  if ($notes) { $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes }
  return $s
}
function Bullets($s, $l, $t, $w, $h, $lines, $size = 16) {
  $b = Txt $s $l $t $w $h ($lines -join "`r") $size $TEXT
  $r = $b.TextFrame.TextRange; $r.ParagraphFormat.Bullet.Visible = -1; $r.ParagraphFormat.Bullet.Character = 8226
  $r.ParagraphFormat.Bullet.Font.Color.RGB = $GOLD; $r.ParagraphFormat.SpaceAfter = 8
  $b.TextFrame.Ruler.Levels.Item(1).FirstMargin = 0; $b.TextFrame.Ruler.Levels.Item(1).LeftMargin = 16
  return $b
}
function Img($s, $path, $l, $t, $maxW, $maxH, $caption) {
  $p = $s.Shapes.AddPicture((Join-Path $root $path), 0, -1, $l, $t)
  $p.LockAspectRatio = -1; $k = [Math]::Min($maxW / $p.Width, $maxH / $p.Height); $p.Width = $p.Width * $k
  $p.Left = $l + ($maxW - $p.Width) / 2; $p.Top = $t
  $p.Line.Visible = -1; $p.Line.ForeColor.RGB = (C 200 205 210)
  if ($caption) { [void](Txt $s $l ($t + $p.Height + 2) $maxW 18 $caption 10 $GRAY -align 2) }
  return $p
}
function Stat($s, $l, $t, $w, $h, $num, $label, $color) {
  [void](Box $s $l $t $w $h $WHITE (C 210 220 230)); [void](Box $s $l $t $w 5 $color $null)
  [void](Txt $s $l ($t + 10) $w 44 $num 30 $color -bold -align 2)
  [void](Txt $s ($l + 6) ($t + 54) ($w - 12) ($h - 58) $label 12 $GRAY -align 2)
}
function Card($s, $l, $t, $w, $h, $head, $body, $color) {
  [void](Box $s $l $t $w $h $WHITE (C 210 220 230)); [void](Box $s $l $t $w 30 $color $null)
  [void](Txt $s ($l + 6) ($t + 4) ($w - 12) 24 $head 14 $WHITE -bold)
  $bt = Txt $s ($l + 8) ($t + 38) ($w - 16) ($h - 42) $body 15 $TEXT; $bt.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 6
}

# ---------------- 1. PORTADA ----------------
$script:n++
$s = $pres.Slides.Add(1, 12)
[void](Box $s 0 0 960 540 $NAVY $null); [void](Box $s 60 292 120 4 $GOLD $null)
[void](Txt $s 60 80 840 30 'BASE DE DATOS AVANZADAS Y BIG DATA · CIIN1487R · NRC 3814' 15 (C 190 210 230) -bold)
[void](Txt $s 60 120 840 110 'DataShield 360: ecosistema de datos seguro e inteligente para la empresa peruana' 34 $WHITE -bold)
[void](Txt $s 60 240 840 36 'Avance semana 6 · Solución de Inteligencia de Negocios y Visualización de Datos' 19 (C 232 200 120))
[void](Txt $s 60 315 840 110 ("Caso: Distribuidora Andina Comercial S.A.C. (Lima, Arequipa y Trujillo)`rIntegrantes: Channel · Yin · Gutiérrez · Pirsin`rDocente: Ramirez Visa Adan Elvis · Período 2026-2") 16 $WHITE)
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = 'Buenos días. Somos el grupo de Channel, Yin, Gutiérrez y Pirsin. Presentamos el avance de la semana 6 del proyecto DataShield 360: llegamos hasta Inteligencia de Negocios y Visualización de Datos. La parte de Big Data y MongoDB es para la entrega final.'

# ---------------- 2. AGENDA ----------------
$s = Slide 'Agenda' 'Qué vamos a presentar' 'Seguimos el mismo orden del informe: capítulos I, II y III.' 'Esta es la ruta de la exposición. Va en el mismo orden del informe: primero el problema, luego cómo diseñamos la solución y al final qué logramos y qué sigue.'
$items = @(
  @('1', 'El problema y los objetivos', 'Capítulo I'),
  @('2', 'Arquitectura y decisiones de diseño', 'Capítulos II y III (3.1 a 3.3)'),
  @('3', 'Base de datos, automatización, ETL y Data Warehouse', 'Capítulo III (3.4.2 y 3.4.3)'),
  @('4', 'Dashboard y lo que nos dice', 'Capítulo III (3.4.4)'),
  @('5', 'Seguridad, prueba de viabilidad y cronograma', 'Capítulo III (3.6 y 3.7)'),
  @('6', 'Conclusiones y próximos pasos', 'Cierre'))
$y = 98
foreach ($it in $items) {
  [void](Box $s 60 $y 46 46 $BLUE $null); [void](Txt $s 60 ($y + 8) 46 30 $it[0] 20 $WHITE -bold -align 2)
  [void](Txt $s 120 ($y + 2) 560 26 $it[1] 18 $NAVY -bold); [void](Txt $s 120 ($y + 25) 560 20 $it[2] 12 $GRAY)
  $y += 56
}

# ---------------- 3. CONTEXTO Y PROBLEMA ----------------
$s = Slide 'Capítulo I · Contexto y problema' 'El problema de Distribuidora Andina' 'La empresa tiene muchos datos, pero no los protege ni los usa para decidir.' 'Distribuidora Andina es una empresa ficticia que vende productos de consumo masivo. Su problema: decide con reportes de Excel hechos a mano a fin de mes, su base de datos no tiene controles y la autoridad le exige cumplir la ley de datos personales en 60 días.'
[void](Bullets $s 40 92 500 340 @('Vende a bodegas, minimarkets y mayoristas desde 6 almacenes', 'Decide con reportes de Excel hechos a mano a fin de mes (1 mes de retraso)', 'Su base SQL Server no tiene procedimientos, auditoría ni respaldos formales', 'Tiene datos con errores: ventas anuladas mezcladas, nombres sucios, emails vacíos', 'Debe cumplir la Ley N.° 29733 de protección de datos en 60 días') 16)
Stat $s 570 96 165 110 '78 %' 'empresas con pérdida o daño de datos (caso)' $RED
Stat $s 755 96 165 110 '91 %' 'medianas empresas deciden sin análisis formal' $RED
Stat $s 570 222 350 110 '200 000 ventas' '3 años de operación (2023-2025) en nuestro dataset de prueba' $BLUE

# ---------------- 4. OBJETIVOS Y ALCANCE ----------------
$s = Slide 'Capítulo I · Objetivos y alcance' 'Qué queremos lograr en este avance' 'Hoy entregamos la solución BI completa; Big Data y NoSQL van en la entrega final.' 'Nuestro objetivo general es tener indicadores de ventas confiables, al menos 5 veces más rápidos y con los datos protegidos. A la izquierda lo que entregamos hoy; a la derecha lo que haremos en la entrega final.'
[void](Txt $s 40 88 880 40 'Objetivo general: indicadores de ventas que cuadren al 100 % con el origen, al menos 5 veces más rápidos que hoy y con los datos personales protegidos.' 15 $NAVY -bold)
Card $s 40 140 430 296 'Incluido hoy (semana 6)' ("• Base de datos de 11 tablas con datos simulados`r• 3 procedimientos y 2 triggers (automatización)`r• 3 roles, auditoría y respaldos`r• ETL con 5 reglas de limpieza`r• Data Warehouse en esquema estrella`r• Dashboard en Excel con filtros OLAP`r• Prueba de viabilidad del rendimiento") $GREEN
Card $s 490 140 430 296 'Entrega final (semanas 7 a 16)' ("• MongoDB: comentarios y rutas en JSON`r• PySpark: análisis con ≥ 10 000 registros`r• Integrar los resultados de Spark al dashboard`r• Experimento completo y conclusiones (capítulos IV a VIII)`r• Informe de Medición") $GOLD

# ---------------- 5. ARQUITECTURA ----------------
$s = Slide 'Capítulo III · 3.3 Arquitectura' 'Cómo viajan los datos' 'Separamos la base del día a día (registrar) del Data Warehouse (analizar).' 'Así viajan los datos: se registran en la base del día a día, los procedimientos y triggers la controlan, el ETL limpia los datos y los pasa al Data Warehouse, y el dashboard los muestra a la gerencia. Abajo, en punteado, lo que agregaremos con Big Data.'
[void](Img $s 'diagramas\E-05c_arquitectura.png' 40 86 880 338 'Figura 1 del informe · Arquitectura integrada de la solución')

# ---------------- 6. DECISIONES DE DISEÑO ----------------
$s = Slide 'Capítulo III · 3.1 y 3.2 Decisiones de diseño' 'Qué elegimos y por qué' 'Cada decisión se comparó con 2 alternativas usando criterios con peso.' 'Para cada parte comparamos tres opciones usando rendimiento, costo, facilidad y seguridad. Elegimos Excel porque ya está instalado y cumple; el esquema estrella porque es simple y rápido; el índice columnar porque la prueba demostró que es unas 10 veces más rápido; y el ETL en SQL porque se puede versionar y probar.'
Card $s 40 92 205 340 'Herramienta BI' ("Elegida: Excel 365`r`rDescartadas: Power BI (no instalado) y Tableau Public (publica los datos en internet)`r`rCosto: S/ 0") $BLUE
Card $s 260 92 205 340 'Modelo del DW' ("Elegido: esquema estrella (Kimball)`r`rDescartados: copo de nieve (más lento) e Inmon (más largo de diseñar)`r`rPocas uniones, fácil de entender") $BLUE
Card $s 480 92 205 340 'Almacenamiento' ("Elegido: índice columnar (COLUMNSTORE)`r`rDescartados: índices normales y vistas indexadas`r`rProbado: 7,7 a 11,5 veces más rápido") $BLUE
Card $s 700 92 220 340 'Proceso ETL' ("Elegido: procedimiento en T-SQL`r`rDescartados: SSIS (herramienta extra) y Python (no instalado)`r`rSe guarda en Git y tiene bitácora") $BLUE

# ---------------- 7. BASE DE DATOS ----------------
$s = Slide 'Capítulo III · 3.4.2 Modelo de datos' 'La base de datos del día a día' 'Las reglas de la base impiden errores antes de que ocurran.' 'Esta es la base transaccional: 11 tablas relacionadas. Cada tabla tiene una llave primaria y las relaciones usan llaves foráneas. Además hay reglas, por ejemplo: el stock nunca puede ser negativo y el RUC no se puede repetir.'
[void](Img $s 'diagramas\E-05a_diagrama_er.png' 40 86 560 338 'Figura 2 · Diagrama entidad-relación')
[void](Bullets $s 620 92 300 340 @('11 tablas relacionadas (clientes, productos, ventas, inventario...)', 'Llaves primarias y foráneas en todas las relaciones', 'Reglas: stock ≥ 0, RUC único, cantidad > 0', 'Datos simulados: 200 000 ventas y 599 806 líneas de detalle', '13 validaciones de calidad: todas PASA') 15)

# ---------------- 8. AUTOMATIZACIÓN ----------------
$s = Slide 'Capítulo III · 3.4.3 Automatización' 'Procedimientos y triggers' 'Una venta se guarda completa o no se guarda: nunca queda a medias.' 'Un procedimiento es una receta guardada que la base ejecuta siempre igual. Un trigger es un vigilante que anota cada cambio. En la captura: la primera venta se guarda; la segunda no tiene stock, así que se deshace todo con ROLLBACK.'
[void](Bullets $s 40 92 400 340 @('usp_RegistrarVenta: venta + detalle + stock en una sola transacción', 'usp_ActualizarInventario: repone o ajusta el stock', 'usp_GenerarAlertasStock: avisa cuando el stock baja del mínimo', '2 triggers anotan quién cambió qué y cuándo (Cliente y Producto)', 'Todos con TRY-CATCH, COMMIT y ROLLBACK') 15)
[void](Img $s 'evidencias\E-06a_procedimiento_ssms.png' 460 90 460 320 'Figura 5 · Venta válida (COMMIT) y venta sin stock (ROLLBACK)')

# ---------------- 9. DW Y ETL ----------------
$s = Slide 'Capítulo III · 3.4.2 y 3.4.3 Data Warehouse y ETL' 'Del dato sucio al dato confiable' 'El total del Data Warehouse cuadra al céntimo con la base original.' 'El ETL extrae los datos, los limpia con 5 reglas y los carga al Data Warehouse, que tiene forma de estrella: en el centro las ventas y alrededor cuándo, qué, quién y dónde. Comprobamos que el total vendido sea igual en ambos lados: diferencia cero.'
[void](Img $s 'diagramas\E-05b_esquema_estrella.png' 40 86 430 300 'Figura 3 · Esquema estrella')
[void](Txt $s 490 88 430 24 '5 reglas de limpieza (ETL)' 16 $NAVY -bold)
[void](Bullets $s 490 114 430 200 @('Nombres en mayúsculas y sin espacios', 'Se excluyen las ventas anuladas', 'Se calculan año, mes, trimestre y campaña', 'No se copian datos personales (RUC, email)', 'Se rechazan cantidades o importes ≤ 0') 14)
Stat $s 490 322 205 112 '582 141' 'líneas cargadas (17 667 anuladas rechazadas)' $BLUE
Stat $s 715 322 205 112 'S/ 0,00' 'diferencia entre la base original y el Data Warehouse' $GREEN

# ---------------- 10. DASHBOARD ----------------
$s = Slide 'Capítulo III · 3.4.4 Dashboard' 'Dashboard de ventas en Excel' 'Con un clic en los filtros se analiza cualquier año, región o tipo de cliente.' 'Este es el dashboard, conectado directamente al Data Warehouse. Uno: filtros. Dos: los indicadores clave. Tres: la tendencia mensual. Cuatro: ventas por región y año. Al tocar un filtro, todo se recalcula: eso es el análisis OLAP.'
[void](Img $s 'evidencias\E-07_dashboard.png' 40 86 880 338 'Figura 7 · Filtros (1), KPI (2), tendencia (3) y análisis por región y año (4)')

# ---------------- 11. QUÉ NOS DICE ----------------
$s = Slide 'Capítulo III · 3.4.4 Interpretación' 'Qué decisiones permite tomar' 'Hay que preparar stock para fin de año y crear acciones para volver a crecer.' 'Lo más importante para la gerencia: noviembre y diciembre venden 2,5 veces más que el resto del año, así que hay que preparar stock desde octubre. Las ventas están estancadas, así que se necesitan nuevas acciones comerciales.'
Stat $s 40 92 205 120 'S/ 180,1 M' 'ventas 2023-2025 (7,28 millones de unidades)' $BLUE
Stat $s 260 92 205 120 '2,5 veces' 'más ventas por mes en noviembre y diciembre' $GOLD
Stat $s 480 92 205 120 '−0,41 %' 'variación de ventas en 2025: están estancadas' $RED
Stat $s 700 92 220 120 '50 / 30 / 20 %' 'Lima / Arequipa / Trujillo' $BLUE
[void](Bullets $s 40 228 470 210 @('Campaña de fin de año → comprar y distribuir stock desde octubre', 'Ventas planas → la meta de crecimiento necesita nuevas promociones', 'Top 10 productos → priorizarlos en stock y promociones', 'Filtro 2025 + Lima → S/ 29,97 millones en un clic') 15)
[void](Img $s 'evidencias\E-07b_dashboard_filtro_lima_2025.png' 530 226 390 210 '')

# ---------------- 12. SEGURIDAD ----------------
$s = Slide 'Capítulo III · 3.6 Seguridad y respaldos' 'Datos protegidos y respaldados' 'Cada persona ve solo lo que necesita y se pierden como máximo 15 minutos de datos.' 'Aplicamos mínimo privilegio: el analista no puede ver email, teléfono ni RUC, como muestra la captura en rojo. Los respaldos completo, diferencial y de log permiten perder como máximo 15 minutos de información.'
Card $s 40 92 280 168 'Roles (mínimo privilegio)' ("• Administrador: control total`r• Analista: solo lectura, sin datos personales`r• Operativo: solo usa los procedimientos, no borra") $BLUE
Card $s 40 270 280 166 'Respaldos (backup)' ("• Completo: domingo 01:00`r• Diferencial: diario 01:00`r• Log: cada 15 minutos`r• Verificados con RESTORE VERIFYONLY") $GREEN
[void](Img $s 'evidencias\E-06c_roles_ssms.png' 340 90 580 300 'Figura 10 · Al analista se le niega el acceso al email del cliente')
[void](Txt $s 340 414 580 22 'Cumple la Ley N.° 29733 e ISO/IEC 27001 (matriz en la sección 3.6 del informe)' 12 $GRAY -align 2)

# ---------------- 13. PRUEBA DE VIABILIDAD ----------------
$s = Slide 'Capítulo III · 3.7 Prueba de viabilidad' '¿El Data Warehouse realmente acelera los reportes?' 'Sí: el Data Warehouse respondió entre 7,7 y 11,5 veces más rápido. Es viable.' 'Era nuestra mayor duda: ¿vale la pena el Data Warehouse o basta con poner índices? Medimos 4 consultas de la gerencia en 3 escenarios, 10 veces cada una. Las barras azules, el Data Warehouse, son siempre las más bajas, es decir, las más rápidas. Por eso el dashboard se conecta al Data Warehouse.'
[void](Img $s 'evidencias\E-09b_grafico_tiempos.png' 40 86 560 320 'Figura 12 · Tiempo promedio por consulta (ms) · menos es mejor')
[void](Bullets $s 620 92 300 330 @('Rojo: base actual', 'Naranja: base + índices', 'Azul: Data Warehouse', '4 consultas × 10 repeticiones', 'Los 3 dan el mismo resultado', 'Los índices solo aceleran buscar un cliente', 'Decisión: el dashboard usa el Data Warehouse') 15)

# ---------------- 14. RIESGOS Y CRONOGRAMA ----------------
$s = Slide 'Capítulo III · 3.6 y 3.7 Riesgos y cronograma' 'Riesgos y plan de trabajo' 'La ruta crítica es PySpark; por eso tiene 3 días extra de contingencia.' 'Los riesgos más altos son instalar Python con PySpark y MongoDB, así que les dimos días extra. Los hitos 1 al 4 ya están completos; lo que sigue es MongoDB, Spark, la integración y el informe final.'
Card $s 40 92 400 344 'Riesgos principales' ("• PySpark no instalado (alta) → instalar o usar Colab`r• MongoDB no instalado (alta) → Community o Atlas`r• Datos simulados con sesgo (media) → 13 validaciones; un sesgo ya fue corregido`r• Tiempos variables (media) → 10 repeticiones`r• Datos personales (baja, alto impacto) → roles y DW sin datos personales") $RED
[void](Txt $s 460 92 460 26 'Cronograma (* = ruta crítica)' 16 $NAVY -bold)
$rows = @(@('Hito','Fechas','Qué se entrega'), @('H1-H4','09/11 – 16/11','BD, seguridad, DW y dashboard ✔'), @('H5','23/11 – 04/12','MongoDB (+2 días)'), @('H6*','07/12 – 23/12','PySpark (+3 días)'), @('H7*','04/01 – 15/01','Integrar Spark al dashboard'), @('H8*','18/01 – 22/01','Informe final y de Medición'), @('H9*','Semana 16','Sustentación final'))
$tb = $s.Shapes.AddTable($rows.Count, 3, 460, 122, 460, 300).Table
$tb.Columns.Item(1).Width = 70; $tb.Columns.Item(2).Width = 130; $tb.Columns.Item(3).Width = 260
for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 3; $j++) {
  $cell = $tb.Cell($i + 1, $j + 1); $tr = $cell.Shape.TextFrame.TextRange; $tr.Text = $rows[$i][$j]; $tr.Font.Name = 'Segoe UI'; $tr.Font.Size = 14
  if ($i -eq 0) { $cell.Shape.Fill.ForeColor.RGB = $NAVY; $tr.Font.Color.RGB = $WHITE; $tr.Font.Bold = -1 }
  else { $cell.Shape.Fill.ForeColor.RGB = $(if ($i -eq 1) { C 226 240 230 } elseif ($i % 2) { $WHITE } else { $LIGHT }); $tr.Font.Color.RGB = $TEXT } } }
# ---------------- 15. CONCLUSIONES ----------------
$s = Slide 'Cierre' 'Conclusiones y próximos pasos' 'La solución BI funciona, está probada y es la base para la parte de Big Data.' 'Para cerrar: el Data Warehouse fue la decisión correcta, el ETL entrega datos confiables, el dashboard muestra decisiones claras y la base ahora es segura. En la entrega final agregaremos MongoDB y Spark sobre esta misma base. Gracias.'
Stat $s 40 92 205 112 '~10 veces' 'reportes más rápidos con el Data Warehouse' $BLUE
Stat $s 260 92 205 112 'S/ 0,00' 'diferencia: el ETL cuadra al céntimo' $GREEN
Stat $s 480 92 205 112 '10 / 10' 'pruebas de funcionamiento y seguridad aprobadas' $GREEN
Stat $s 700 92 220 112 '2,5 veces' 'más ventas en la campaña de fin de año' $GOLD
[void](Bullets $s 40 222 430 214 @('Separar la base del día a día del Data Warehouse fue la decisión correcta', 'Los índices normales no resuelven los reportes', 'El dashboard muestra decisiones claras para la gerencia', 'La base es segura: transacciones, auditoría, roles y respaldos') 15)
Card $s 490 222 430 214 'Próximos pasos' ("• MongoDB con operaciones CRUD`r• PySpark: ranking y estacionalidad`r• Integrar Spark al dashboard`r• Experimento completo e Informe de Medición`r• Publicar el repositorio para el docente") $GOLD

$out = Join-Path $root 'informe\AVANCE-EXP-DEDATOSAVANZADO.pptx'
if (Test-Path $out) { Remove-Item $out -Force }
$pres.SaveAs($out)
$pres.SaveAs((Join-Path $root 'informe\AVANCE-EXP-DEDATOSAVANZADO_presentacion.pdf'), 32)
"Diapositivas: $($pres.Slides.Count) -> $out"
$pres.Close()

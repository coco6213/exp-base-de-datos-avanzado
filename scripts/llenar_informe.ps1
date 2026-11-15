# Llena la plantilla del Informe de Proyecto Final con informe\contenido.txt usando Microsoft Word (COM).
# Requisito: informe\trabajo.docx (copia limpia de la plantilla) abierto en Microsoft Word (no WPS).
# Formato de contenido.txt:
#   @@H <titulo>   reemplaza el texto guia bajo ese titulo
#   @@S <inicio>   reemplaza el parrafo que empieza con ese texto (sin lineas = lo borra)
#   @@T <n>        llena la tabla n de la plantilla (campos separados por |)
#   @@NT <clave>   tabla nueva que se inserta donde aparezca [[NT:clave]]
#   [[IMG:ruta|leyenda|ancho_cm]] inserta imagen;  lineas "CODE: ..." = codigo en Consolas
param([string]$Salida = 'Grupo_CIIN1487R_Avance_S6')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$doc = [Runtime.InteropServices.Marshal]::BindToMoniker("$root\informe\trabajo.docx")
$word = $doc.Application
$word.ScreenUpdating = $false
function U([int[]]$codes) { -join ($codes | % { [char]$_ }) }   # texto con tildes sin depender de la codificacion del script

# ---------- leer contenido ----------
$lines = Get-Content "$root\informe\contenido.txt" -Encoding UTF8
$blocks = New-Object System.Collections.ArrayList
$cur = $null
foreach ($l in $lines) {
  if ($l -match '^@@(\w+)\s+(.+)$') { $cur = [pscustomobject]@{Kind=$Matches[1]; Key=$Matches[2].Trim(); Lines=(New-Object System.Collections.ArrayList)}; [void]$blocks.Add($cur) }
  elseif ($cur -ne $null) { if ($l.Trim() -ne '') { [void]$cur.Lines.Add($l) } }
}
function Block($kind, $key) { $blocks | ? { $_.Kind -eq $kind -and $_.Key -eq $key } | Select-Object -First 1 }

# ---------- utilidades ----------
$T = @{}; for ($i = 1; $i -le $doc.Tables.Count; $i++) { $T[$i] = $doc.Tables.Item($i) }
function SetCell($tbl, $r, $c, $txt) { $tbl.Cell($r, $c).Range.Text = $txt }
function FillTable($tbl, $rows) {
  $need = 1 + $rows.Count
  while ($tbl.Rows.Count -lt $need) { [void]$tbl.Rows.Add() }
  while ($tbl.Rows.Count -gt $need) { $tbl.Rows.Item($tbl.Rows.Count).Delete() }
  for ($r = 0; $r -lt $rows.Count; $r++) {
    $f = $rows[$r].Split('|')
    for ($c = 0; $c -lt $f.Count -and $c -lt $tbl.Columns.Count; $c++) { SetCell $tbl ($r + 2) ($c + 1) $f[$c].Trim() }
  }
  $tbl.Range.Font.Italic = $false
}
function IsHeading($p) { return ($p.Style.NameLocal -match '^T.tulo') }
function FindPara($text) {
  $rng = $doc.Content
  $f = $rng.Find; $f.ClearFormatting(); $f.Text = $text; $f.MatchCase = $true; $f.Forward = $true; $f.Wrap = 0
  if ($f.Execute()) { return $rng.Paragraphs.Item(1) } else { return $null }
}
function FindHeading($text) {
  $rng = $doc.Content
  while ($true) {
    $f = $rng.Find; $f.ClearFormatting(); $f.Text = $text; $f.MatchCase = $true; $f.Forward = $true; $f.Wrap = 0
    if (-not $f.Execute()) { return $null }
    $p = $rng.Paragraphs.Item(1)
    if ((IsHeading $p) -and ($p.Range.Text.Trim() -eq $text)) { return $p }
    $rng.SetRange($rng.End, $doc.Content.End)
  }
}
function FormatBody($start, $len) {
  $r = $doc.Range($start, $start + $len + 1)
  $r.Style = $doc.Styles.Item('Normal')
  $r.ListFormat.RemoveNumbers()
  $r.Font.Reset()
  $r.ParagraphFormat.Alignment = 3   # justificado
}
function WriteInto($para, $textLines) {
  # escribe el texto dentro de un parrafo existente (sin tocar su marca final)
  $txt = ($textLines -join "`r")
  $start = $para.Range.Start
  $doc.Range($start, $para.Range.End - 1).Text = $txt
  FormatBody $start $txt.Length
}
function InsertAfter($para, $textLines) {
  # agrega el texto al final del titulo y lo separa con un salto: nunca cae dentro de una tabla siguiente
  $txt = ($textLines -join "`r")
  $pos = $para.Range.End - 1
  $doc.Range($pos, $pos).InsertAfter("`r" + $txt)
  FormatBody ($pos + 1) $txt.Length
}function ReplaceAll($from, $to, $rangeObj) {
  $rng = $rangeObj.Duplicate
  $f = $rng.Find; $f.ClearFormatting(); $f.Text = $from; $f.MatchCase = $true; $f.Forward = $true; $f.Wrap = 0
  $n = 0
  while ($f.Execute()) { $rng.Text = $to; $rng.Collapse(0); $n++; $f = $rng.Find; $f.Text = $from; $f.Forward = $true; $f.Wrap = 0 }
  return $n
}

# ---------- 1) tablas existentes ----------
foreach ($b in ($blocks | ? Kind -eq 'T')) { FillTable $T[[int]$b.Key] $b.Lines; Write-Host "Tabla $($b.Key) llena ($($b.Lines.Count) filas)" }

# ---------- 2) portada, encabezado y pie ----------
$gut = 'GUTI' + (U 0xC9) + 'RREZ'
$nombres = @('CHANNEL', 'YIN', $gut, 'PIRSIN')
for ($i = 0; $i -lt 4; $i++) { SetCell $T[1] ($i + 2) 2 $nombres[$i]; SetCell $T[1] ($i + 2) 3 ('25 % de participaci' + (U 0xF3) + 'n') }
SetCell $T[1] 8 2 ('Repositorio Git local, etiqueta v0.6-avance-s6 (URL pendiente de publicaci' + (U 0xF3) + 'n)')
$titulo = 'DataShield 360 ' + (U 0x2013) + ' Distribuidora Andina Comercial S.A.C.'
$autores = 'Channel; Guti' + (U 0xE9) + 'rrez; Pirsin; Yin'
foreach ($s in $doc.Sections) {
  foreach ($k in 1..3) {
    [void](ReplaceAll ('<<T' + (U 0xED) + 'tulo del proyecto>>') $titulo $s.Headers.Item($k).Range)
    $rng = $s.Footers.Item($k).Range.Duplicate; $f = $rng.Find; $f.ClearFormatting()
    $f.Text = '\<\<Colocar*coma\>^t\>'; $f.MatchWildcards = $true; $f.Forward = $true; $f.Wrap = 0
    if ($f.Execute()) { $rng.Text = $autores }
  }
}
# ---------- 3) declaracion IA (casillas) y firmas ----------
$box = U 0x2610; $chk = U 0x2612
$marcar = @(('S' + (U 0xED)), 'Claude', 'Ideas', ('Redacci' + (U 0xF3) + 'n'), ('C' + (U 0xF3) + 'digo'), 'Errores', ('Gr' + (U 0xE1) + 'ficos'), ('An' + (U 0xE1) + 'lisis'), 'Diagrama', 'Resultados', 'Conclusiones', 'Reporte', ('Ejecut' + (U 0xE9) + ' el c' + (U 0xF3) + 'digo'), ('Compar' + (U 0xE9) + ' resultados'), ('Correg' + (U 0xED) + ' errores'))
$ia = $T[24].Range
foreach ($m in $marcar) { $n = ReplaceAll "$box $m" "$chk $m" $ia; Write-Host "Casilla '$m': $n" }
$t25 = $T[25]
for ($r = 2; $r -le $t25.Rows.Count; $r++) {
  $c1 = ($t25.Cell($r, 1).Range.Text -replace '[\r\a]', '').Trim()
  if ($c1 -like 'Fecha*') { SetCell $t25 $r 2 '15/11/2026' }
  elseif ($c1 -eq '' -and ($r - 2) -lt 4) { SetCell $t25 $r 1 $nombres[$r - 2] }
}

# ---------- 4) secciones por titulo ----------
foreach ($b in ($blocks | ? Kind -eq 'H')) {
  $h = FindHeading $b.Key
  if ($h -eq $null) { Write-Warning "No se encontro el titulo: $($b.Key)"; continue }
  $g = New-Object System.Collections.ArrayList
  $n = $h.Next(1)
  while ($n -ne $null -and -not (IsHeading $n) -and -not $n.Range.Information(12)) { [void]$g.Add($n); $n = $n.Next(1) }
  if ($g.Count -eq 0) { if ($b.Lines.Count -gt 0) { InsertAfter $h $b.Lines } }
  else {
    for ($k = $g.Count - 1; $k -ge 1; $k--) { $g[$k].Range.Delete() | Out-Null }
    if ($b.Lines.Count -gt 0) { WriteInto $g[0] $b.Lines } else { $g[0].Range.Delete() | Out-Null }
  }
  Write-Host "Seccion: $($b.Key)"
}

# ---------- 5) reemplazo de parrafos especificos ----------
foreach ($b in ($blocks | ? Kind -eq 'S')) {
  $p = FindPara $b.Key
  if ($p -eq $null) { Write-Warning "No se encontro: $($b.Key)"; continue }
  if ($b.Lines.Count -eq 0) { $p.Range.Delete() | Out-Null; continue }
  $r = $doc.Range($p.Range.Start, $p.Range.End - 1); $r.Text = ($b.Lines -join "`r")
  $r.Font.Reset(); $r.Font.Italic = $false
}
$abet = FindPara 'ABET. (2026)'; $abadi = FindPara 'Abadi, D. J.'
if ($abet -ne $null -and $abadi -ne $null) { $abadi.Format = $abet.Format; $abadi.Range.Font.Reset() }

# ---------- 6) marcadores: tablas nuevas e imagenes ----------
while ($true) {
  $p = FindPara '[['
  if ($p -eq $null) { break }
  $tag = $p.Range.Text.Trim()
  if ($tag -match '^\[\[NT:(\w+)\]\]$') {
    $b = Block 'NT' $Matches[1]
    $rows = @($b.Lines | % { ($_.Split('|') | % { $_.Trim() }) -join "`t" })
    $r = $doc.Range($p.Range.Start, $p.Range.End - 1); $r.Text = ($rows -join "`r")
    $cols = $b.Lines[0].Split('|').Count
    $tb = $r.ConvertToTable(1, $rows.Count, $cols)
    $tb.Style = 'Tabla con cuadr' + (U 0xED) + 'cula'
    $tb.Range.Font.Size = 8.5; $tb.Range.ParagraphFormat.Alignment = 0; $tb.Range.ParagraphFormat.SpaceAfter = 0
    $tb.Rows.Item(1).Range.Font.Bold = $true; $tb.Rows.Item(1).Shading.BackgroundPatternColor = 16117480
    $tb.Rows.Item(1).Range.Font.Color = 4531467; $tb.Rows.Item(1).HeadingFormat = $true
    $tb.AutoFitBehavior(1); $tb.AutoFitBehavior(2)
    $prev = $tb.Range.Paragraphs.Item(1).Previous(1); if ($prev -ne $null) { $prev.KeepWithNext = -1 }
    Write-Host "Tabla nueva: $($Matches[1])"
  }
  elseif ($tag -match '^\[\[IMG:(.+)\]\]$') {
    $parts = $Matches[1].Split('|'); $img = Join-Path $root $parts[0]; $cap = $parts[1]; $w = [double]$parts[2]
    $r = $doc.Range($p.Range.Start, $p.Range.End - 1); $r.Text = ''
    $pic = $doc.InlineShapes.AddPicture($img, $false, $true, $r)
    $ratio = $pic.Height / $pic.Width; $pic.Width = $w * 28.35; $pic.Height = $pic.Width * $ratio
    $p.Alignment = 1; $p.SpaceAfter = 2; $p.KeepWithNext = -1
    $pos = $p.Range.End - 1; $doc.Range($pos, $pos).InsertAfter("`r")   # divide el parrafo: la leyenda nunca cae en una tabla
    $cp = $doc.Range($pos + 1, $pos + 1).Paragraphs.Item(1)
    $doc.Range($cp.Range.Start, $cp.Range.End - 1).Text = $cap
    $cp = $doc.Range($pos + 1, $pos + 1).Paragraphs.Item(1)
    $cp.Style = $doc.Styles.Item('Normal'); $cp.Range.Font.Reset(); $cp.Range.Font.Size = 9; $cp.Range.Font.Italic = $true; $cp.Alignment = 1; $cp.KeepWithNext = 0
    Write-Host "Imagen: $($parts[0])"
  }
  else { Write-Warning "Marcador desconocido: $tag"; break }
}

# ---------- 7) lineas de codigo ----------
while ($true) {
  $p = FindPara 'CODE: '
  if ($p -eq $null) { break }
  $doc.Range($p.Range.Start, $p.Range.Start + 6).Delete() | Out-Null
  $p.Range.Font.Name = 'Consolas'; $p.Range.Font.Size = 8.5; $p.Alignment = 0
  $p.Shading.BackgroundPatternColor = 15921906
}

# ---------- 8) indice y guardado ----------
foreach ($lv in 1..3) { $doc.Styles.Item("Titulo $lv").ParagraphFormat.OutlineLevel = $lv }
if ($doc.TablesOfContents.Count -gt 0) { $doc.TablesOfContents.Item(1).Update() }
$word.ScreenUpdating = $true
$docx = "$root\informe\$Salida.docx"; $pdf = "$root\informe\$Salida.pdf"
$doc.SaveAs2($docx, 16)
$doc.ExportAsFixedFormat($pdf, 17)
"Paginas: " + $doc.ComputeStatistics(2)
"Guardado: $docx y $pdf"

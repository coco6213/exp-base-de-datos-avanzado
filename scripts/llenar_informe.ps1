# Llena la plantilla del Informe de Proyecto Final con informe\contenido.txt usando Microsoft Word (COM).
# Requisito: el documento informe\trabajo.docx abierto en Microsoft Word (no WPS).
param([string]$Salida = 'Grupo_CIIN1487R_Avance_S6')
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$doc = [Runtime.InteropServices.Marshal]::BindToMoniker("$root\informe\trabajo.docx")
$word = $doc.Application
$word.ScreenUpdating = $false

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
function InsertAfter($para, $textLines) {
  # inserta parrafos normales despues de $para
  $txt = ($textLines -join "`r")
  $r = $para.Range
  $r.InsertParagraphAfter()
  $np = $para.Next(1)
  $np.Range.Text = $txt
  $np = $para.Next(1)
  $start = $np.Range.Start
  $endR = $doc.Range($start, $start + $txt.Length + 1)
  $endR.Style = $doc.Styles.Item('Normal')
  $endR.Font.Reset()
  $endR.ParagraphFormat.Alignment = 3   # justificado
}

# ---------- 1) tablas existentes ----------
foreach ($b in ($blocks | ? Kind -eq 'T')) { FillTable $T[[int]$b.Key] $b.Lines; Write-Host "Tabla $($b.Key) llena ($($b.Lines.Count) filas)" }

# ---------- 2) portada (T1) ----------
$nombres = @('CHANNEL', 'YIN', 'GUTIERREZ', 'PIRSIN')
$nombres[2] = 'GUTI' + [char]0x00C9 + 'RREZ'
for ($i = 0; $i -lt 4; $i++) { SetCell $T[1] ($i + 2) 2 $nombres[$i]; SetCell $T[1] ($i + 2) 3 ('25 % de participaci' + [char]0x00F3 + 'n') }
SetCell $T[1] 8 2 ('Repositorio Git local, etiqueta v0.6-avance-s6 (URL pendiente de publicaci' + [char]0x00F3 + 'n)')

# ---------- 3) declaracion IA (casillas) y firmas ----------
$box = [string][char]0x2610; $chk = [string][char]0x2612
$marcar = @('S' + [char]0x00ED, 'Claude', 'Ideas', 'Redacci' + [char]0x00F3 + 'n', 'C' + [char]0x00F3 + 'digo', 'Errores', 'Gr' + [char]0x00E1 + 'ficos', 'An' + [char]0x00E1 + 'lisis', 'Diagrama', 'Resultados', 'Conclusiones', 'Reporte', 'Ejecut' + [char]0x00E9 + ' el c' + [char]0x00F3 + 'digo', 'Compar' + [char]0x00E9 + ' resultados', 'Correg' + [char]0x00ED + ' errores')
foreach ($m in $marcar) {
  $rng = $doc.Content; $f = $rng.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
  [void]$f.Execute("$box $m", $true, $false, $false, $false, $false, $true, 0, $false, "$chk $m", 2)
}
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
  # borrar la guia (parrafos hasta el siguiente titulo o tabla)
  $n = $h.Next(1)
  while ($n -ne $null -and -not (IsHeading $n) -and -not $n.Range.Information(12)) {
    $nx = $n.Next(1); $n.Range.Delete() | Out-Null; $n = $nx
  }
  if ($b.Lines.Count -gt 0) { InsertAfter $h $b.Lines }
  Write-Host "Seccion: $($b.Key)"
}

# ---------- 5) reemplazo de parrafos especificos ----------
foreach ($b in ($blocks | ? Kind -eq 'S')) {
  $p = FindPara $b.Key
  if ($p -eq $null) { Write-Warning "No se encontro: $($b.Key)"; continue }
  $r = $doc.Range($p.Range.Start, $p.Range.End - 1); $r.Text = ($b.Lines -join "`r")
  $r.Font.Reset(); $r.Font.Italic = $false
}

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
    $tb.Style = 'Tabla con cuadr' + [char]0x00ED + 'cula'
    $tb.Range.Font.Size = 8.5; $tb.Range.ParagraphFormat.Alignment = 0; $tb.Range.ParagraphFormat.SpaceAfter = 0
    $tb.Rows.Item(1).Range.Font.Bold = $true; $tb.Rows.Item(1).Shading.BackgroundPatternColor = 16117480
    $tb.Rows.Item(1).Range.Font.Color = 4531467; $tb.Rows.Item(1).HeadingFormat = $true
    $tb.AutoFitBehavior(1); $tb.AutoFitBehavior(2)
    Write-Host "Tabla nueva: $($Matches[1])"
  }
  elseif ($tag -match '^\[\[IMG:(.+)\]\]$') {
    $parts = $Matches[1].Split('|'); $img = Join-Path $root $parts[0]; $cap = $parts[1]; $w = [double]$parts[2]
    $r = $doc.Range($p.Range.Start, $p.Range.End - 1); $r.Text = ''
    $pic = $doc.InlineShapes.AddPicture($img, $false, $true, $r)
    $ratio = $pic.Height / $pic.Width; $pic.Width = $w * 28.35; $pic.Height = $pic.Width * $ratio
    $p.Alignment = 1; $p.SpaceAfter = 2
    $p.Range.InsertParagraphAfter(); $cp = $p.Next(1)
    $cp.Range.Text = $cap
    $cp = $p.Next(1); $cp.Range.Font.Reset(); $cp.Range.Font.Size = 9; $cp.Range.Font.Italic = $true; $cp.Alignment = 1
    Write-Host "Imagen: $($parts[0])"
  }
  else { Write-Warning "Marcador desconocido: $tag"; break }
}

# ---------- 7) indice y guardado ----------
if ($doc.TablesOfContents.Count -gt 0) { $doc.TablesOfContents.Item(1).Update() }
$word.ScreenUpdating = $true
$docx = "$root\informe\$Salida.docx"; $pdf = "$root\informe\$Salida.pdf"
$doc.SaveAs2($docx, 16)
$doc.ExportAsFixedFormat($pdf, 17)
"Paginas: " + $doc.ComputeStatistics(2)
"Guardado: $docx y $pdf"

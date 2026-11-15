# Genera los graficos del experimento (tiempo y lecturas por escenario) con Excel y los exporta a PNG.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$q = "SET NOCOUNT ON; SELECT ConsultaID, Escenario, CAST(AVG(TiempoMs) AS DECIMAL(10,1)) T, CAST(STDEV(TiempoMs) AS DECIMAL(10,1)) D, AVG(Lecturas) L FROM dbo.Resultado WHERE TipoCorrida='CALIENTE' GROUP BY ConsultaID, Escenario ORDER BY ConsultaID, Escenario"
$rows = sqlcmd -S localhost -E -C -d LabExperimento -h -1 -W -s "|" -Q $q | ? { $_ -match '\|' } | % { $c = $_.Split('|'); [pscustomobject]@{Q=$c[0];E=$c[1];T=[double]$c[2];D=[double]$c[3];L=[double]$c[4]} }
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$wb = $xl.Workbooks.Add(); $ws = $wb.Worksheets.Item(1); $ws.Name = 'Experimento'
$nom = @{Q1='Q1 Region x mes'; Q2='Q2 Top 10 productos'; Q3='Q3 Categoria x anio'; Q4='Q4 Cliente puntual'}
$ws.Range('A1:D1').Value2 = [object[]]@('Consulta','E1 OLTP base','E2 OLTP + indices','E3 DW columnstore')
$ws.Range('F1:I1').Value2 = [object[]]@('Consulta','E1 OLTP base','E2 OLTP + indices','E3 DW columnstore')
$fila = @{Q1=2;Q2=3;Q3=4;Q4=5}; $colT = @{E1='B';E2='C';E3='D'}; $colL = @{E1='G';E2='H';E3='I'}
foreach ($k in $fila.Keys) { [string]$a = 'A' + $fila[$k]; [string]$b = 'F' + $fila[$k]; $ws.Range($a).Value2 = $nom[$k]; $ws.Range($b).Value2 = $nom[$k] }
foreach ($x in $rows) {
  [string]$ct = $colT[$x.E] + $fila[$x.Q]; [string]$cl = $colL[$x.E] + $fila[$x.Q]
  $ws.Range($ct).Value2 = [double]$x.T; $ws.Range($cl).Value2 = [double]$x.L
}
function G($rango, $titulo, $eje, $file, $top) {
  $sh = $ws.Shapes.AddChart2(-1, 51, 20, $top, 760, 380)
  $ch = $sh.Chart; $ch.SetSourceData($ws.Range($rango)); $ch.PlotBy = 2
  $ch.HasTitle = $true; $ch.ChartTitle.Text = $titulo
  $ch.Axes(2).HasTitle = $true; $ch.Axes(2).AxisTitle.Text = $eje
  $ch.HasLegend = $true; $ch.Legend.Position = -4107
  $ch.SeriesCollection(1).Format.Fill.ForeColor.RGB = 0x5A5AC0   # rojo suave (BGR)
  $ch.SeriesCollection(2).Format.Fill.ForeColor.RGB = 0x2E9BE8   # naranja
  $ch.SeriesCollection(3).Format.Fill.ForeColor.RGB = 0x8B5A1F   # azul
  for ($s = 1; $s -le 3; $s++) { $ch.SeriesCollection($s).HasDataLabels = $true; $ch.SeriesCollection($s).DataLabels().Font.Size = 9 }
  Start-Sleep 1; $ch.Export($file) | Out-Null
}
G 'A1:D5' 'Tiempo promedio por consulta (ms, 10 corridas en caliente) - menos es mejor' 'Milisegundos' "$root\evidencias\E-09b_grafico_tiempos.png" 160
G 'F1:I5' 'Paginas leidas por consulta (lecturas logicas) - menos es mejor' 'Paginas de 8 KB' "$root\evidencias\E-09c_grafico_lecturas.png" 560
$wb.SaveAs("$root\resultados\RUN-03_graficos_experimento.xlsx", 51); $wb.Close($false)
"OK"

# Convierte paginas de un PDF a PNG usando la API de Windows (Windows.Data.Pdf). Sirve para revisar el informe.
param([string]$Pdf, [string]$Dir, [int]$Desde = 1, [int]$Hasta = 999, [int]$Ancho = 1100)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
$asTaskAction = ([System.WindowsRuntimeSystemExtensions].GetMethods() | ? { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(); $task.Result }
function AwaitA($op) { $asTaskAction.Invoke($null, @($op)).Wait() }
$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Pdf)) ([Windows.Storage.StorageFile])
$pdfDoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
New-Item -ItemType Directory -Force $Dir | Out-Null
$n = [Math]::Min($Hasta, $pdfDoc.PageCount)
for ($i = $Desde; $i -le $n; $i++) {
  $page = $pdfDoc.GetPage($i - 1)
  $opt = New-Object Windows.Data.Pdf.PdfPageRenderOptions; $opt.DestinationWidth = $Ancho
  $ms = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
  AwaitA ($page.RenderToStreamAsync($ms, $opt))
  $out = Join-Path $Dir ('p{0:D2}.png' -f $i)
  $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($ms.GetInputStreamAt(0))
  $fs = [IO.File]::Create($out); $net.CopyTo($fs); $fs.Close(); $ms.Dispose(); $page.Dispose()
}
"Paginas totales: $($pdfDoc.PageCount); renderizadas $Desde-$n en $Dir"

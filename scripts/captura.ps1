# Captura de pantalla de una ventana (por su proceso) y marcado con cuadros rojos.
#   Captura: captura.ps1 -Proceso EXCEL -Salida img.png
#   Marcar : captura.ps1 -Marcar img.png -Cuadros "x,y,w,h;x,y,w,h" [-Etiquetas "1;2"]
param([string]$Proceso, [string]$Salida, [string]$Marcar, [string]$Cuadros, [string]$Etiquetas, [switch]$Pantalla)
Add-Type -AssemblyName System.Drawing
Add-Type @"
using System; using System.Runtime.InteropServices;
public class W32 {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  public struct RECT { public int L, T, R, B; }
}
"@
[W32]::SetProcessDPIAware() | Out-Null

if ($Marcar) {
  $img = [Drawing.Image]::FromFile($Marcar); $bmp = New-Object Drawing.Bitmap $img; $img.Dispose()
  $g = [Drawing.Graphics]::FromImage($bmp); $g.SmoothingMode = 'AntiAlias'
  $pen = New-Object Drawing.Pen ([Drawing.Color]::Red), 5
  $font = New-Object Drawing.Font 'Segoe UI', 16, ([Drawing.FontStyle]::Bold)
  $et = if ($Etiquetas) { $Etiquetas.Split(';') } else { @() }; $i = 0
  foreach ($c in $Cuadros.Split(';')) {
    $v = $c.Split(',') | % { [int]$_ }
    $g.DrawRectangle($pen, $v[0], $v[1], $v[2], $v[3])
    if ($i -lt $et.Count -and $et[$i]) {
      $sz = $g.MeasureString($et[$i], $font)
      $y = [Math]::Max(0, $v[1] - $sz.Height - 2)
      $g.FillRectangle([Drawing.Brushes]::Red, $v[0], $y, $sz.Width + 6, $sz.Height)
      $g.DrawString($et[$i], $font, [Drawing.Brushes]::White, $v[0] + 3, $y)
    }
    $i++
  }
  $g.Dispose(); $bmp.Save($Marcar + '.tmp.png', [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
  Move-Item -Force ($Marcar + '.tmp.png') $Marcar
  "Marcado: $Marcar"; return
}

$p = Get-Process $Proceso | ? { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
[W32]::keybd_event(0x12, 0, 0, 0); [W32]::keybd_event(0x12, 0, 2, 0)   # ALT: permite traer la ventana al frente
[W32]::ShowWindow($p.MainWindowHandle, 3) | Out-Null                    # maximizar
[W32]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
Start-Sleep -Milliseconds 300
[W32]::keybd_event(0x1B, 0, 0, 0); [W32]::keybd_event(0x1B, 0, 2, 0)   # ESC: oculta las letras de atajo que muestra ALT
[W32]::SetCursorPos(1900, 600) | Out-Null                              # aparta el mouse (evita tooltips)
Start-Sleep -Milliseconds 1500
$r = New-Object W32+RECT; [W32]::GetWindowRect($p.MainWindowHandle, [ref]$r) | Out-Null
$L = [Math]::Max(0, $r.L); $T = [Math]::Max(0, $r.T)
Add-Type -AssemblyName System.Windows.Forms
$b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$W = [Math]::Min($r.R, $b.Width) - $L; $H = [Math]::Min($r.B, $b.Height) - $T
$bmp = New-Object Drawing.Bitmap $W, $H
$g = [Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen($L, $T, 0, 0, $bmp.Size); $g.Dispose()
$bmp.Save($Salida, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
"Captura: $Salida ($W x $H)"

# Ayudante de interfaz: clic, teclas y captura de pantalla completa.
#  ui.ps1 -Clic "x,y"   |  ui.ps1 -Teclas "{F5}" -Proceso SSMS  |  ui.ps1 -Captura ruta.png
param([string]$Clic, [string]$Teclas, [string]$Proceso, [string]$Captura)
Add-Type -AssemblyName System.Drawing, System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public class U32 {
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(int f, int x, int y, int d, int e);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, int f, int e);
}
"@
[U32]::SetProcessDPIAware() | Out-Null
if ($Proceso) {
  $p = Get-Process $Proceso | ? { $_.MainWindowHandle -ne 0 } | Select-Object -First 1
  if ($p) { [U32]::keybd_event(0x12,0,0,0); [U32]::keybd_event(0x12,0,2,0)
            [U32]::ShowWindow($p.MainWindowHandle, 3) | Out-Null; [U32]::SetForegroundWindow($p.MainWindowHandle) | Out-Null
            Start-Sleep -Milliseconds 400; [U32]::keybd_event(0x1B,0,0,0); [U32]::keybd_event(0x1B,0,2,0); Start-Sleep -Milliseconds 400 }
}
if ($Clic) { $v = $Clic.Split(',') | % { [int]$_ }; [U32]::SetCursorPos($v[0], $v[1]) | Out-Null; Start-Sleep -Milliseconds 150
             [U32]::mouse_event(2,0,0,0,0); [U32]::mouse_event(4,0,0,0,0); Start-Sleep -Milliseconds 300 }
if ($Teclas) { [System.Windows.Forms.SendKeys]::SendWait($Teclas); Start-Sleep -Milliseconds 300 }
if ($Captura) {
  [U32]::SetCursorPos(1915, 500) | Out-Null; Start-Sleep -Milliseconds 600
  $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $bmp = New-Object Drawing.Bitmap $b.Width, ($b.Height - 60)      # sin la barra de tareas
  $g = [Drawing.Graphics]::FromImage($bmp); $g.CopyFromScreen(0, 0, 0, 0, $bmp.Size); $g.Dispose()
  $bmp.Save($Captura, [Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose(); "Captura: $Captura"
}

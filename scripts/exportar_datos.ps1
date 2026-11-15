# Exporta las tablas de SQL Server a archivos CSV (separados por ;) en la carpeta DATOS, para abrirlos con Excel.
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot
$base = Join-Path $root 'DATOS_FICTICIOS'
$plan = @(
  @{Db='DistribuidoraAndina'; Dir='1_Base_dia_a_dia (OLTP)'; Tablas=@('Region','Sede','Categoria','Producto','Cliente','Vendedor','Inventario','Venta','DetalleVenta','AuditoriaLog','AlertaStock')},
  @{Db='DW_Andina';           Dir='2_Data_Warehouse (DW)';   Tablas=@('DimTiempo','DimProducto','DimCliente','DimSede','FactVentas','ETL_Log','vw_Dashboard')},
  @{Db='LabExperimento';      Dir='3_Prueba_de_viabilidad';  Tablas=@('Consulta','Resultado','Almacenamiento')}
)
$utf8 = New-Object Text.UTF8Encoding $true   # con BOM: Excel muestra bien las tildes
foreach ($p in $plan) {
  $dir = Join-Path $base $p.Dir; New-Item -ItemType Directory -Force $dir | Out-Null
  $cn = New-Object Data.SqlClient.SqlConnection "Server=localhost;Database=$($p.Db);Integrated Security=True;TrustServerCertificate=True"
  $cn.Open()
  foreach ($t in $p.Tablas) {
    $cmd = $cn.CreateCommand(); $cmd.CommandText = "SELECT * FROM dbo.[$t]"; $cmd.CommandTimeout = 300
    $rd = $cmd.ExecuteReader()
    $file = Join-Path $dir "$t.csv"
    $sw = New-Object IO.StreamWriter($file, $false, $utf8)
    $cols = 0..($rd.FieldCount - 1)
    $sw.WriteLine(($cols | % { $rd.GetName($_) }) -join ';')
    $n = 0
    while ($rd.Read()) {
      $vals = foreach ($i in $cols) {
        if ($rd.IsDBNull($i)) { '' } else {
          $v = $rd.GetValue($i)
          if ($v -is [datetime]) { $v.ToString('yyyy-MM-dd HH:mm:ss') }
          elseif ($v -is [decimal] -or $v -is [double]) { $v.ToString([Globalization.CultureInfo]::InvariantCulture) }
          else { $s = [string]$v; if ($s -match '[;"\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s } }
        }
      }
      $sw.WriteLine($vals -join ';'); $n++
    }
    $sw.Close(); $rd.Close()
    "{0,-22} {1,-18} {2,9:N0} filas" -f $p.Dir.Split(' ')[0], $t, $n
  }
  $cn.Close()
}


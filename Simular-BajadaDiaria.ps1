<#
.SYNOPSIS
    Simula la bajada diaria de vulnerabilidades de infraestructura sobre un archivo de nombre fijo:
    Reporte_Vulnerabilidades_Infra.xlsx

.DESCRIPTION
    Cada ejecución representa un día de trabajo:
      - Si el archivo no existe, lo crea con 8 vulnerabilidades iniciales.
      - Si existe, lo lee y:
          * agrega de 1 a 3 vulnerabilidades nuevas con un estado de ingreso: Nueva, Abierta o Rechazada,
          * vuelve a abrir como "Resurgente" alguna cerrada que el escáner detecta de nuevo,
          * hace avanzar el estado de algunas existentes (las Rechazadas se revisan y pueden
            terminar en "Cancelado", con el motivo del rechazo en Comentarios),
          * mueve el comentario anterior a "Comentarios histórico" con su fecha,
          * actualiza Asignado y Fecha estimada.
      - Sobrescribe SIEMPRE el mismo archivo.

    Es reproducible: la misma fecha sobre el mismo archivo produce los mismos cambios.
    Requiere el módulo ImportExcel (Install-Module ImportExcel -Scope CurrentUser). No necesita Excel instalado.

.PARAMETER Carpeta
    Carpeta donde vive el reporte. Por defecto, la carpeta del script.

.PARAMETER Fecha
    Día que se simula (formato AAAA-MM-DD). Por defecto, hoy.

.PARAMETER Reiniciar
    Borra el reporte y empieza desde cero.

.PARAMETER Forzar
    Permite procesar de nuevo una fecha que ya figura en el reporte.

.EXAMPLE
    .\Simular-BajadaDiaria.ps1 -Reiniciar -Fecha 2026-10-07
    .\Simular-BajadaDiaria.ps1 -Fecha 2026-10-08
    .\Simular-BajadaDiaria.ps1
#>
[CmdletBinding()]
param(
    [string]$Carpeta = $PSScriptRoot,
    [datetime]$Fecha = (Get-Date).Date,
    [switch]$Reiniciar,
    [switch]$Forzar
)

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel
Add-Type -AssemblyName System.Drawing

$NombreArchivo = 'Reporte_Vulnerabilidades_Infra.xlsx'
$Hoja          = 'Vulnerabilidades'
$Columnas      = 'Id', 'Nombre', 'Descripción', 'Estado', 'Comentarios', 'Comentarios histórico', 'Asignado', 'Fecha estimada'
$Ruta          = Join-Path $Carpeta $NombreArchivo
$hoy           = $Fecha.Date
$fechaTxt      = $hoy.ToString('yyyy-MM-dd')
# Semilla derivada de un hash: con semillas consecutivas, System.Random da secuencias casi iguales día a día.
$hash          = [System.Security.Cryptography.MD5]::Create().ComputeHash([Text.Encoding]::UTF8.GetBytes($fechaTxt))
$rnd           = [System.Random]::new([BitConverter]::ToInt32($hash, 0))

#region Datos de ejemplo
$Catalogo = @(
    @{ Nombre = 'OpenSSH desactualizado';                    Equipo = 'Linux';         Descripcion = 'La versión de OpenSSH tiene vulnerabilidades conocidas (CVE-2024-6387 "regreSSHion") que permiten ejecución remota de código sin autenticación.' }
    @{ Nombre = 'Kernel Linux sin parches';                  Equipo = 'Linux';         Descripcion = 'Kernel vulnerable a escalada local de privilegios (CVE-2024-1086, use-after-free en netfilter).' }
    @{ Nombre = 'Sudo vulnerable (Baron Samedit)';           Equipo = 'Linux';         Descripcion = 'Versión de sudo afectada por CVE-2021-3156: un usuario local puede obtener privilegios de root.' }
    @{ Nombre = 'SMBv1 habilitado';                          Equipo = 'Windows';       Descripcion = 'El protocolo SMBv1 está activo. Es obsoleto y explotable (EternalBlue, MS17-010); debe deshabilitarse.' }
    @{ Nombre = 'Parches de seguridad de Windows pendientes'; Equipo = 'Windows';      Descripcion = 'Faltan actualizaciones acumulativas de seguridad de los últimos 2 meses, incluidas correcciones críticas.' }
    @{ Nombre = 'RDP expuesto sin NLA';                      Equipo = 'Windows';       Descripcion = 'Escritorio remoto accesible sin Network Level Authentication, lo que facilita fuerza bruta y explotación previa a la autenticación.' }
    @{ Nombre = 'Print Spooler habilitado en servidor';      Equipo = 'Windows';       Descripcion = 'El servicio de cola de impresión está activo en un servidor que no imprime (riesgo PrintNightmare, CVE-2021-34527).' }
    @{ Nombre = 'LLMNR / NetBIOS habilitado';                Equipo = 'Windows';       Descripcion = 'La resolución de nombres por LLMNR/NBT-NS permite capturar hashes NTLM mediante envenenamiento (Responder).' }
    @{ Nombre = 'Credenciales por defecto en equipo de red'; Equipo = 'Redes';         Descripcion = 'El equipo acepta el usuario y la contraseña de fábrica en la interfaz de administración.' }
    @{ Nombre = 'SNMP v1/v2c con comunidad pública';         Equipo = 'Redes';         Descripcion = 'SNMP responde con la comunidad "public", lo que expone la configuración y la topología de la red.' }
    @{ Nombre = 'Telnet habilitado';                         Equipo = 'Redes';         Descripcion = 'La administración por Telnet transmite credenciales en texto plano. Debe reemplazarse por SSH.' }
    @{ Nombre = 'Firmware de firewall desactualizado';       Equipo = 'Redes';         Descripcion = 'La versión de firmware tiene una vulnerabilidad crítica publicada por el fabricante en la VPN SSL.' }
    @{ Nombre = 'Apache Log4j vulnerable (Log4Shell)';       Equipo = 'Middleware';    Descripcion = 'La aplicación incluye log4j-core 2.14, afectada por CVE-2021-44228 (ejecución remota de código).' }
    @{ Nombre = 'TLS 1.0 / 1.1 habilitado';                  Equipo = 'Middleware';    Descripcion = 'El servicio acepta protocolos TLS obsoletos y suites de cifrado débiles.' }
    @{ Nombre = 'Certificado SSL vencido';                   Equipo = 'Middleware';    Descripcion = 'El certificado del servicio HTTPS está vencido, lo que genera alertas y facilita ataques de intermediario.' }
    @{ Nombre = 'Apache Tomcat desactualizado';              Equipo = 'Middleware';    Descripcion = 'La versión de Tomcat está fuera de soporte y tiene múltiples CVE publicadas.' }
    @{ Nombre = 'Cabeceras de seguridad HTTP ausentes';      Equipo = 'Middleware';    Descripcion = 'Faltan las cabeceras HSTS, X-Content-Type-Options y Content-Security-Policy.' }
    @{ Nombre = 'Cuenta sa con contraseña débil';            Equipo = 'Base de datos'; Descripcion = 'La cuenta administradora del motor SQL tiene una contraseña trivial que se obtuvo por diccionario.' }
    @{ Nombre = 'Puerto de base de datos expuesto';          Equipo = 'Base de datos'; Descripcion = 'El puerto del motor de base de datos es accesible desde toda la red corporativa, sin filtrado.' }
    @{ Nombre = 'Motor de base de datos sin soporte';        Equipo = 'Base de datos'; Descripcion = 'La versión del motor de base de datos llegó al fin de soporte y ya no recibe parches de seguridad.' }
)

$Activos = @{
    'Linux'         = 'srv-lnx-web01', 'srv-lnx-app02', 'srv-lnx-bat03'
    'Windows'       = 'srv-win-ad01', 'srv-win-file02', 'srv-win-rds03'
    'Redes'         = 'sw-core-01', 'fw-perim-01', 'rt-wan-02'
    'Middleware'    = 'srv-web-dmz01', 'srv-tomcat-02', 'srv-api-03'
    'Base de datos' = 'srv-sql-01', 'srv-ora-02'
}

$Responsables = @{
    'Linux'         = 'Equipo Linux - M. Gómez'
    'Windows'       = 'Equipo Windows - L. Fernández'
    'Redes'         = 'Equipo Redes - P. Sosa'
    'Middleware'    = 'Equipo Middleware - C. Ruiz'
    'Base de datos' = 'Equipo DBA - A. Medina'
}

$MotivosRechazo = @(
    'el activo fue dado de baja y ya no está en producción'
    'el servicio no está expuesto: el puerto está filtrado por el firewall'
    'la versión instalada no coincide con la que informa el escáner'
    'falso positivo del escáner: la detección se basa solo en el banner del servicio'
    'el componente vulnerable no está en uso por la aplicación'
)

$ColorEstado = @{
    'Nueva'                = '#FCE4D6'
    'Abierta'              = '#FFC7CE'
    'Resurgente'           = '#FF9999'
    'Rechazada'            = '#E4C1F9'
    'Cancelado'            = '#BFBFBF'
    'En análisis'          = '#FFEB9C'
    'En remediación'       = '#F8CBAD'
    'Pendiente validación' = '#BDD7EE'
    'Cerrada'              = '#C6EFCE'
    'Riesgo aceptado'      = '#D9D9D9'
}

# Estados que quedan fuera del cálculo de penalidad. Rechazada NO está: sigue contando hasta que se revise.
$EstadosSinPenalidad = 'Cerrada', 'Riesgo aceptado', 'Cancelado'

$EquipoPorNombre = @{}
foreach ($c in $Catalogo) { $EquipoPorNombre[$c.Nombre] = $c.Equipo }
#endregion

#region Funciones auxiliares
function Get-Azar($lista) { $lista[$rnd.Next($lista.Count)] }

function ConvertTo-Fecha($valor) {
    if ($null -eq $valor -or "$valor" -eq '') { return $null }
    if ($valor -is [datetime]) { return $valor.Date }
    if ($valor -is [double] -or $valor -is [int]) { return [datetime]::FromOADate([double]$valor).Date }
    return [datetime]::ParseExact("$valor", 'yyyy-MM-dd', $null).Date
}

# El comentario vigente pasa al histórico (el más nuevo arriba) y se escribe el del día.
function Set-Comentario($fila, [string]$texto) {
    if ($fila.Comentarios) {
        $fila.'Comentarios histórico' = if ($fila.'Comentarios histórico') {
            "$($fila.Comentarios)`n$($fila.'Comentarios histórico')"
        } else { $fila.Comentarios }
    }
    $fila.Comentarios = "[$fechaTxt] $texto"
}

function Get-Equipo($fila) { $EquipoPorNombre[($fila.Nombre -replace '\s\([^)]*\)$', '')] }

function New-Vulnerabilidad {
    $disponibles = @(foreach ($c in $Catalogo) {
        foreach ($a in $Activos[$c.Equipo]) {
            $n = "$($c.Nombre) ($a)"
            if (-not $nombresUsados.Contains($n)) { [pscustomobject]@{ Nombre = $n; Descripcion = $c.Descripcion } }
        }
    })
    if ($disponibles.Count -eq 0) { return $null }

    $elegida = Get-Azar $disponibles
    [void]$nombresUsados.Add($elegida.Nombre)
    $script:ultimoId++

    $fila = [pscustomobject][ordered]@{
        'Id'                    = 'VULN-{0:D4}' -f $script:ultimoId
        'Nombre'                = $elegida.Nombre
        'Descripción'           = $elegida.Descripcion
        'Estado'                = $null
        'Comentarios'           = $null
        'Comentarios histórico' = $null
        'Asignado'              = 'Sin asignar'
        'Fecha estimada'        = $hoy.AddDays((Get-Azar 15, 30, 45))
    }

    # Estado de ingreso: mayormente Nueva; a veces ya figura como Abierta o el responsable la rechaza.
    $p = $rnd.NextDouble()
    if ($p -lt 0.60) {
        $fila.Estado = 'Nueva'
        Set-Comentario $fila 'Primera detección en el escaneo diario de infraestructura. Pendiente de asignación.'
    } elseif ($p -lt 0.85) {
        $fila.Estado = 'Abierta'
        Set-Comentario $fila 'Detectada en el escaneo diario de infraestructura. Pendiente de asignación.'
    } else {
        Set-Rechazo $fila
    }
    $ingreso[$fila.Estado]++
    $fila
}

# El equipo responsable rechaza el hallazgo. El motivo queda en Comentarios para que Seguridad lo revise.
function Set-Rechazo($fila) {
    $responsable = $Responsables[(Get-Equipo $fila)]
    $fila.Estado = 'Rechazada'
    $fila.Asignado = $responsable
    Set-Comentario $fila "Rechazada por $responsable. Motivo: $(Get-Azar $MotivosRechazo). Seguridad debe revisarla."
}

function Get-MotivoRechazo($fila) {
    $m = [regex]::Match("$($fila.Comentarios)`n$($fila.'Comentarios histórico')", 'Motivo: ([^.]+)')
    if ($m.Success) { $m.Groups[1].Value } else { 'no informado' }
}
#endregion

#region Carga del reporte existente
if ($Reiniciar -and (Test-Path $Ruta)) { Remove-Item $Ruta }

$filas = [System.Collections.Generic.List[object]]::new()
if (Test-Path $Ruta) {
    foreach ($f in Import-Excel -Path $Ruta -WorksheetName $Hoja) {
        $filas.Add([pscustomobject][ordered]@{
            'Id'                    = $f.Id
            'Nombre'                = $f.Nombre
            'Descripción'           = $f.'Descripción'
            'Estado'                = $f.Estado
            'Comentarios'           = $f.Comentarios
            'Comentarios histórico' = $f.'Comentarios histórico'
            'Asignado'              = $f.Asignado
            'Fecha estimada'        = (ConvertTo-Fecha $f.'Fecha estimada')
        })
    }
}

if (-not $Forzar -and ($filas | Where-Object { "$($_.Comentarios)".StartsWith("[$fechaTxt]") })) {
    Write-Warning "El reporte ya tiene novedades del $fechaTxt. Usá otra -Fecha, o -Forzar para procesarla de nuevo."
    return
}

$nombresUsados = [System.Collections.Generic.HashSet[string]]::new()
$ultimoId = 0
foreach ($f in $filas) {
    [void]$nombresUsados.Add($f.Nombre)
    $num = [int]($f.Id -replace '\D', '')
    if ($num -gt $ultimoId) { $ultimoId = $num }
}
#endregion

#region Novedades del día
$esPrimerDia = $filas.Count -eq 0
$conCambios = 0
$cerradas = 0
$canceladas = 0
$ingreso = [ordered]@{ 'Nueva' = 0; 'Abierta' = 0; 'Resurgente' = 0; 'Rechazada' = 0 }

foreach ($fila in $filas) {
    $antes = $fila.Comentarios
    $equipo = Get-Equipo $fila
    $p = $rnd.NextDouble()

    switch ($fila.Estado) {
        'Nueva' {
            if ($p -lt 0.5) {
                $fila.Estado = 'En análisis'
                $fila.Asignado = $Responsables[$equipo]
                Set-Comentario $fila "Asignada a $($Responsables[$equipo]). Se valida el alcance y el impacto en el activo."
            } else {
                $fila.Estado = 'Abierta'
                Set-Comentario $fila 'Sigue sin asignar. Pasa a Abierta.'
            }
        }
        'Abierta' {
            if ($p -lt 0.6) {
                $fila.Estado = 'En análisis'
                $fila.Asignado = $Responsables[$equipo]
                Set-Comentario $fila "Asignada a $($Responsables[$equipo]). Se valida el alcance y el impacto en el activo."
            }
        }
        'Resurgente' {
            if ($p -lt 0.6) {
                $fila.Estado = 'En análisis'
                $fila.Asignado = $Responsables[$equipo]
                Set-Comentario $fila "Reincidencia asignada a $($Responsables[$equipo]). Se analiza por qué no se sostuvo la corrección anterior."
            }
        }
        'Rechazada' {
            if ($p -lt 0.45) {
                $fila.Estado = 'En análisis'
                $fila.Asignado = $Responsables[$equipo]
                Set-Comentario $fila "Seguridad revisó el rechazo y no se sostiene: la vulnerabilidad aplica. Se reasigna a $($Responsables[$equipo])."
            } elseif ($p -lt 0.80) {
                $fila.Estado = 'Cancelado'
                Set-Comentario $fila "Cancelado. Motivo de rechazo validado por Seguridad: $(Get-MotivoRechazo $fila)."
                $canceladas++
            }
        }
        'Cerrada' {
            if ($p -lt 0.08) {
                $fila.Estado = 'Resurgente'
                $fila.'Fecha estimada' = $hoy.AddDays(15)
                Set-Comentario $fila 'El escaneo vuelve a detectar la vulnerabilidad (reincidencia). Se reabre.'
                $ingreso['Resurgente']++
            }
        }
        'En análisis' {
            if ($p -lt 0.10) {
                $fila.Estado = 'Riesgo aceptado'
                $fila.'Fecha estimada' = $hoy.AddDays(90)
                Set-Comentario $fila "El dueño del activo acepta el riesgo con un control compensatorio documentado. Se revisa el $($hoy.AddDays(90).ToString('yyyy-MM-dd'))."
            } elseif ($p -lt 0.55) {
                $fila.Estado = 'En remediación'
                $fila.'Fecha estimada' = $hoy.AddDays((Get-Azar 7, 10, 14))
                Set-Comentario $fila ("Cambio CHG-{0} aprobado. La corrección se aplica en la próxima ventana de mantenimiento." -f $rnd.Next(10000, 99999))
            } elseif ($p -lt 0.65) {
                Set-Rechazo $fila
            }
        }
        'En remediación' {
            if ($p -lt 0.45) {
                $fila.Estado = 'Pendiente validación'
                $fila.'Fecha estimada' = $hoy.AddDays(2)
                Set-Comentario $fila 'Corrección aplicada. Se solicita un nuevo escaneo para validarla.'
            } elseif ($p -lt 0.60) {
                $fila.'Fecha estimada' = ([datetime]$fila.'Fecha estimada').AddDays(7)
                Set-Comentario $fila 'Se reprograma la corrección porque no hay ventana de mantenimiento disponible.'
            }
        }
        'Pendiente validación' {
            if ($p -lt 0.70) {
                $fila.Estado = 'Cerrada'
                Set-Comentario $fila 'El nuevo escaneo no detecta la vulnerabilidad. Se cierra.'
                $cerradas++
            } elseif ($p -lt 0.85) {
                $fila.Estado = 'En remediación'
                $fila.'Fecha estimada' = $hoy.AddDays(7)
                Set-Comentario $fila 'El nuevo escaneo sigue detectando la vulnerabilidad. Se reabre la remediación.'
            }
        }
    }

    # Sin novedades y con la fecha estimada vencida: se escala.
    if ($fila.Comentarios -eq $antes -and $fila.Estado -notin $EstadosSinPenalidad -and
        $fila.'Fecha estimada' -and [datetime]$fila.'Fecha estimada' -lt $hoy) {
        $fila.'Fecha estimada' = $hoy.AddDays(7)
        Set-Comentario $fila 'La fecha estimada venció. Se escala al responsable y se fija una nueva fecha.'
    }

    if ($fila.Comentarios -ne $antes) { $conCambios++ }
}

$cantidadNuevas = if ($esPrimerDia) { 8 } else { $rnd.Next(1, 4) }
$nuevas = 0
for ($i = 0; $i -lt $cantidadNuevas; $i++) {
    $nueva = New-Vulnerabilidad
    if ($nueva) { $filas.Add($nueva); $nuevas++ }
}
#endregion

#region Escritura sobre el mismo archivo
try {
    $pkg = $filas | Select-Object $Columnas |
        Export-Excel -Path $Ruta -WorksheetName $Hoja -ClearSheet -TableName 'TablaVulnerabilidades' `
            -TableStyle Medium2 -FreezeTopRow -AutoSize -PassThru
} catch {
    throw "No se pudo escribir '$Ruta'. Si está abierto en Excel, cerralo y volvé a ejecutar. Detalle: $($_.Exception.Message)"
}

$ws = $pkg.Workbook.Worksheets[$Hoja]
$ultimaFila = $filas.Count + 1
foreach ($col in @{ 2 = 45; 3 = 60; 5 = 60; 6 = 80; 7 = 30; 8 = 15 }.GetEnumerator()) {
    $ws.Column($col.Key).Width = $col.Value
    $ws.Column($col.Key).Style.WrapText = $true
}
$ws.Column(8).Style.Numberformat.Format = 'yyyy-mm-dd'
$ws.Cells["A1:H$ultimaFila"].Style.VerticalAlignment = 'Top'

for ($r = 2; $r -le $ultimaFila; $r++) {
    $color = $ColorEstado["$($ws.Cells["D$r"].Value)"]
    if ($color) {
        $ws.Cells["D$r"].Style.Fill.PatternType = 'Solid'
        $ws.Cells["D$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml($color))
    }
}
Close-ExcelPackage $pkg
#endregion

#region Resumen en consola
$pendientes = @($filas | Where-Object { $_.Estado -notin $EstadosSinPenalidad }).Count
Write-Host ""
Write-Host "Bajada del $fechaTxt -> $Ruta" -ForegroundColor Cyan
Write-Host ("  Ingreso: Nueva {0} | Abierta {1} | Resurgente {2} | Rechazada {3}" -f $ingreso['Nueva'], $ingreso['Abierta'], $ingreso['Resurgente'], $ingreso['Rechazada'])
Write-Host ("  Altas: {0} | Con novedades: {1} | Cerradas hoy: {2} | Canceladas hoy: {3} | Pendientes (cuentan para penalidad): {4} | Total: {5}" -f $nuevas, $conCambios, $cerradas, $canceladas, $pendientes, $filas.Count)
$filas | Group-Object Estado | Sort-Object Name | Format-Table @{ n = 'Estado'; e = { $_.Name } }, Count -AutoSize
#endregion

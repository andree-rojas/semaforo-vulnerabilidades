# Prompt: Reporte diario de vulnerabilidades de infraestructura

Archivo de trabajo (nombre fijo): **`Reporte_Vulnerabilidades_Infra.xlsx`**
Columnas clave: **Id · Nombre · Descripción · Estado · Comentarios · Comentarios histórico · Asignado · Fecha estimada**

---

## 1. Prompt básico (para crear el Excel de ejemplo)

Copiá y pegá en Claude (claude.ai o Claude Code):

```text
Actuá como analista de seguridad informática de infraestructura.

Contexto:
Todos los días bajamos de nuestro escáner la información de las vulnerabilidades a corregir
y la volcamos SIEMPRE sobre el mismo archivo, de nombre fijo "Reporte_Vulnerabilidades_Infra.xlsx".

Tarea:
Generá un Excel de ejemplo con ese nombre y con una hoja "Vulnerabilidades" que tenga
exactamente estas columnas, en este orden:
Id | Nombre | Descripción | Estado | Comentarios | Comentarios histórico | Asignado | Fecha estimada

Datos:
- Entre 10 y 15 vulnerabilidades realistas de infraestructura (servidores Linux/Windows,
  equipos de red, middleware, bases de datos). Por ejemplo: OpenSSH desactualizado,
  SMBv1 habilitado, TLS 1.0 activo, certificado vencido, Log4j, credenciales por defecto.
- Id con el formato VULN-0001, único e incremental.
- Estado: Abierta, En análisis, En remediación, Pendiente validación, Cerrada o Riesgo aceptado.
- Asignado: el equipo responsable (Linux, Windows, Redes, Middleware, DBA) o "Sin asignar".
- Fecha estimada con el formato AAAA-MM-DD.

Reglas de la actualización diaria (simulá al menos 3 días sobre el mismo archivo):
1. El Id es la clave: nunca se duplica ni se borra una fila.
2. Las vulnerabilidades nuevas entran con el Estado "Abierta".
3. Si cambia el Estado, se reemplaza el valor anterior.
4. "Comentarios" tiene solo el comentario del día, con el formato "[AAAA-MM-DD] texto".
5. Antes de escribir un comentario nuevo, el anterior se MUEVE a "Comentarios histórico",
   con el más reciente primero y uno por línea.
6. Las vulnerabilidades cerradas se conservan para auditoría.

Formato de salida:
El archivo .xlsx con la tabla con filtros, la fila de encabezado fija y el ajuste de texto
en las columnas de comentarios. Agregá también un resumen breve: nuevas, con novedades y
cerradas en el último día.
```

## 2. Variante: actualizar el reporte del día (en claude.ai)

Adjuntá el `Reporte_Vulnerabilidades_Infra.xlsx` de ayer y la bajada nueva del escáner (CSV o Excel), y pegá esto:

```text
Adjunto el Reporte_Vulnerabilidades_Infra.xlsx de ayer y la bajada de hoy del escáner.
Actualizá el reporte para la fecha [AAAA-MM-DD] y respetá estas reglas:
- Hacé coincidir las filas por Id. Si una vulnerabilidad de la bajada no existe en el reporte,
  agregala con el Estado "Abierta" y el siguiente Id libre.
- Si una vulnerabilidad del reporte no aparece en la bajada, pasala a
  "Pendiente validación" (no la borres).
- Por cada fila que cambie, mové el comentario actual a "Comentarios histórico"
  (el más reciente primero) y escribí el nuevo comentario como "[AAAA-MM-DD] texto".
- No modifiques Asignado ni Fecha estimada salvo que la bajada lo indique.
- Devolvé el archivo con el MISMO nombre y las mismas columnas en el mismo orden,
  y un resumen: nuevas, modificadas y cerradas.
```

---

## 3. Qué modelo elegir

| Modelo | Cuándo usarlo para este caso |
|---|---|
| **Claude Sonnet 5.5 (recomendado)** | Generar el Excel de ejemplo, actualizar el reporte diario y escribir o ajustar el script. Razona bien, es rápido y por API cuesta la mitad que Opus (en claude.ai consume menos del límite de uso). Es el punto de equilibrio para una tarea estructurada y repetitiva. |
| Claude Haiku 5.5 | Tareas mecánicas y de mucho volumen con el formato ya definido: resumir el día, clasificar o normalizar textos de la bajada. Es el más barato y el más rápido. |
| Claude Opus 5.5 | Análisis de fondo: priorizar por riesgo (CVSS, exposición, criticidad del activo), correlacionar con CVE y redactar planes de remediación o informes para la gerencia. |
| Claude Fable 5.1 | El más capaz y el más caro. Para este caso no hace falta: reservalo para investigaciones muy complejas o para procesos largos y autónomos. |

**Regla práctica:** empezá con **Sonnet 5.5**. Si notás errores al aplicar las reglas (por ejemplo, comentarios que no pasan bien al histórico), subí a **Opus 5.5**. Si es solo un resumen repetitivo, bajá a **Haiku 5.5**.

En claude.ai el modelo se elige en el selector junto al cuadro del mensaje. En Claude Code se cambia con `/model`.

---

## 4. Simulador local (sin IA): `Simular-BajadaDiaria.ps1`

Para demostrar la bajada diaria sin depender de un modelo, el script de PowerShell actualiza el mismo `Reporte_Vulnerabilidades_Infra.xlsx` en cada ejecución.

```powershell
# Una sola vez: instalar el módulo (no requiere Excel ni permisos de administrador)
Install-Module ImportExcel -Scope CurrentUser

# Empezar de cero y simular varios días
.\Simular-BajadaDiaria.ps1 -Reiniciar -Fecha 2026-10-07
.\Simular-BajadaDiaria.ps1 -Fecha 2026-10-08
.\Simular-BajadaDiaria.ps1 -Fecha 2026-10-09

# Bajada de hoy
.\Simular-BajadaDiaria.ps1
```

- Cada ejecución equivale a un día:
  - Agrega de 1 a 3 vulnerabilidades con un estado de ingreso (Nueva, Abierta o Rechazada) y reabre
    como **Resurgente** alguna cerrada que el escáner vuelve a detectar.
  - Hace avanzar los estados (→ En análisis → En remediación → Pendiente validación → Cerrada, a veces
    Riesgo aceptado).
  - Las **Rechazadas** se revisan: o vuelven a En análisis, o pasan a **Cancelado** con el motivo del
    rechazo en Comentarios. Solo Cerrada, Riesgo aceptado y Cancelado quedan fuera de la penalidad.
  - Mueve el comentario anterior al histórico y escala las que tienen la fecha estimada vencida.
- Si ya se procesó esa fecha, avisa y no hace nada (`-Forzar` lo permite igual).
- La misma fecha produce siempre los mismos cambios, así la demostración es reproducible.
- Si el Excel está abierto, el script no puede escribirlo: cerralo y volvé a ejecutar.

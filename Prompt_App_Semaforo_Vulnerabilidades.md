# Prompt: aplicación web con semáforo de vulnerabilidades

Entrada: el **`Reporte_Vulnerabilidades_Infra.xlsx`** generado por `Simular-BajadaDiaria.ps1`
Columnas: **Id · Nombre · Descripción · Estado · Comentarios · Comentarios histórico · Asignado · Fecha estimada**

> **Ojo:** el Excel **no tiene una columna de severidad**. La aplicación la deduce siempre del Nombre y la
> Descripción con reglas por palabra clave. La columna **Estado** indica cómo llega cada vulnerabilidad
> nueva en la bajada: **Abierta, Nueva, Resurgente o Rechazada**. Solo el estado **Cancelado**, que debe
> traer el motivo del rechazo en Comentarios, excluye una vulnerabilidad del cálculo de penalidad.

---

## Prompt de ejemplo

Copiá y pegá en Claude (claude.ai o Claude Code):

```text
Actuá como analista de seguridad informática y desarrollador web.

OBJETIVO
Construí una aplicación web sencilla para subir el archivo "Reporte_Vulnerabilidades_Infra.xlsx"
y ver un resumen de vulnerabilidades priorizado por severidad y por el riesgo de recibir
penalidades en el corto plazo (por incumplir los plazos de remediación). La visualización
principal es un SEMÁFORO que muestre la prioridad real, no solo la severidad técnica.

ARCHIVO DE ENTRADA
Una hoja "Vulnerabilidades" con estas columnas:
Id | Nombre | Descripción | Estado | Comentarios | Comentarios histórico | Asignado | Fecha estimada
- Estado:
  * Una vulnerabilidad nueva SIEMPRE llega con uno de estos estados de ingreso:
    Abierta    = detectada y pendiente de gestión.
    Nueva      = primera detección en la bajada de hoy.
    Resurgente = ya había sido cerrada y el escáner la vuelve a detectar (reincidencia).
    Rechazada  = el equipo responsable rechazó el hallazgo. Seguridad tiene que volver a revisarlo
                 con prioridad alta y, mientras tanto, SIGUE contando para la penalidad.
  * Después avanza por los estados de gestión: En análisis, En remediación, Pendiente validación,
    Cerrada, Riesgo aceptado o Cancelado.
  * Cancelado = Seguridad confirmó el rechazo (por ejemplo, un falso positivo o un activo dado de
    baja). Se excluye del cálculo de penalidad y SIEMPRE tiene el motivo del rechazo en
    "Comentarios". Si un Cancelado no tiene motivo, no se excluye: tratalo como "Rechazada" y
    mostrá la advertencia "Cancelado sin motivo".
  * Comparalos sin distinguir mayúsculas ni acentos. Si llega un estado desconocido, tratalo
    como "Abierta" y mostrá una advertencia.
- Los comentarios tienen el formato "[AAAA-MM-DD] texto". El histórico tiene una entrada por línea,
  la más reciente primero.
- La Fecha estimada puede venir como fecha de Excel o como texto AAAA-MM-DD.

REGLAS DE NEGOCIO
1. Severidad (Crítica / Alta / Media / Baja):
   - Dedúcela siempre de Nombre y Descripción con una tabla de reglas por palabra clave que quede
     en un objeto de configuración fácil de editar:
     Crítica: Log4j, regreSSHion / OpenSSH, SMBv1, credenciales por defecto, firmware de firewall,
              contraseña débil en cuenta administradora.
     Alta:    kernel, sudo, RDP, Print Spooler, parches de Windows pendientes, Telnet,
              versión sin soporte, Tomcat desactualizado, puerto de base de datos expuesto.
     Media:   TLS 1.0/1.1, certificado vencido, SNMP, LLMNR/NetBIOS.
     Baja:    cabeceras HTTP y todo lo que no coincida con otra regla.
   - En el detalle de cada fila, mostrá qué palabra clave determinó la severidad.

2. Plazo máximo de remediación (SLA) por severidad, también configurable:
   Crítica 7 días, Alta 30 días, Media 90 días, Baja 180 días.
   - Fecha de detección = la fecha más antigua que aparezca en "Comentarios histórico" o en
     "Comentarios". Si el Estado es "Resurgente", el plazo se reinicia: se usa la fecha del
     comentario actual.
   - Vencimiento del SLA = fecha de detección + plazo de su severidad.
   - Días restantes = vencimiento del SLA - hoy. Si el número es negativo, el SLA está vencido.

3. Semáforo (prioridad real). Se evalúa en este orden:
   ROJO (actuar ya):
     - el Estado es "Rechazada" (hay que volver a revisar el hallazgo ya: si el rechazo no se
       sostiene, el plazo sigue corriendo), o
     - el Estado es "Resurgente" y la severidad es Crítica o Alta (la reincidencia indica que la
       corrección anterior no se sostuvo), o
     - el SLA está vencido o la Fecha estimada ya pasó, o
     - la severidad es Crítica o Alta y el SLA vence en 7 días o menos, o
     - la Fecha estimada es posterior al vencimiento del SLA (la penalidad es segura si no se adelanta).
   AMARILLO (atención):
     - el Estado es "Resurgente" con severidad Media o Baja, o
     - el Estado es "Nueva" o "Abierta" y la severidad es Crítica (hay que asignarla y analizarla
       de inmediato), o
     - el SLA vence en 30 días o menos, o
     - la severidad es Crítica o Alta y el Asignado es "Sin asignar", o
     - el estado volvió de "Pendiente validación" a "En remediación" (la corrección falló).
   VERDE (bajo control): el resto de las pendientes, que tienen margen.
   GRIS (fuera del cálculo de penalidad): Cerrada, Riesgo aceptado y Cancelado (con motivo). Se
     muestran aparte, con la fecha de revisión del riesgo aceptado y el motivo de la cancelación
     (tomado de "Comentarios").

4. Orden de prioridad dentro de cada color: primero la severidad (Crítica > Alta > Media > Baja),
   después los días restantes de SLA (los menos primero) y por último el Id.

PANTALLAS
- Zona de carga del archivo (arrastrar y soltar o botón). Solo se aceptan archivos .xlsx.
- Tarjetas con los totales por color del semáforo, con colores grandes y bien visibles, más el
  total de "penalidades probables en los próximos 7 días" y el ingreso del día: cuántas llegaron
  como Abierta, Nueva, Resurgente y Rechazada, y cuántas se cancelaron.
- Tabla priorizada con estas columnas: Semáforo | Id | Nombre | Severidad | Estado | Asignado |
  Fecha estimada | Vencimiento del SLA | Días restantes | Último comentario.
  Al expandir una fila se ven la Descripción y el histórico completo.
- Filtros por color, severidad, estado y Asignado, y un buscador por texto.
- Un gráfico de barras apiladas por Asignado (equipo) con la cantidad de cada color.
- Un botón para exportar a CSV la vista priorizada.
- Una leyenda que explique en una línea el criterio de cada color.

REQUISITOS TÉCNICOS Y DE SEGURIDAD
- Un único archivo index.html, sin backend. El Excel se procesa SOLO en el navegador y no se envía
  a ningún servidor, porque contiene información sensible de vulnerabilidades.
- Leé el .xlsx con SheetJS cargado desde un CDN con la versión fija.
- Escapá todo el texto que venga del Excel antes de mostrarlo (nada de innerHTML con datos del
  archivo), para evitar XSS desde celdas manipuladas.
- Agregá una política Content-Security-Policy en una etiqueta <meta> y un límite de tamaño de
  archivo de 10 MB.
- Validá que estén las columnas obligatorias y mostrá un mensaje claro si falta alguna.
- Tiene que verse bien en pantallas de escritorio y de notebook. Los colores del semáforo deben
  acompañarse con un ícono o texto (no depender solo del color, por accesibilidad).
- El código tiene que tener comentarios en español, y las reglas de severidad, los SLA y los
  umbrales tienen que estar agrupados al principio del script.

ENTREGABLES
1. El archivo index.html completo.
2. Instrucciones para desplegarlo: abrirlo localmente y servirlo con un servidor estático
   (por ejemplo, IIS o un contenedor nginx en la intranet).
3. Una breve explicación de cómo ajustar los SLA y las reglas de severidad.
```

---

## Qué modelo usar

| Modelo | Por qué |
|---|---|
| **Claude Sonnet 5.5 (recomendado)** | Alcanza de sobra para una aplicación de un solo archivo con reglas claras. Es rápido y económico. |
| Claude Opus 5.5 | Elegilo si después querés que la aplicación crezca (backend, usuarios o historial de cargas) o si Sonnet no aplica bien las reglas del semáforo. |
| Claude Haiku 5.5 | No es el indicado para generar la aplicación. Sirve después para tareas chicas, como resumir el CSV exportado. |

**Tip:** en claude.ai, la aplicación se puede ver y probar en el panel de vista previa antes de desplegarla.
Para las pruebas, subí el Excel de ejemplo. Tiene 12 días de historia, vulnerabilidades cerradas
y aceptadas, y fechas estimadas variadas.

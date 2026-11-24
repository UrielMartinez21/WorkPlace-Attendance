# Control de asistencia con Google Chat y Google Sheets (`v9.js`)

Script de Google Apps Script que lee los mensajes de saludo de un espacio de Google Chat, decide quién registró su entrada y quién no, y deja el resultado en una hoja de Google Sheets. Opcionalmente envía un correo con quienes llegaron tarde o no reportaron.

Este documento describe la versión **`v9.js`**, la última de la carpeta.

---

## 1. Qué hace

- Importa los miembros del espacio de Chat a una lista de empleados.
- Busca los mensajes de un día y detecta cuáles son un registro de asistencia (por ejemplo "Buen día", "Buenos días, reporto mi entrada", "Presente").
- Conserva el **primer** mensaje válido de cada persona en el día.
- Marca a cada empleado como `PRESENT` (a tiempo o tarde), `NO CHECK-IN` o `ON VACATION`.
- Guarda el historial en la pestaña `Attendance` (con filtro y colores).
- Rellena una pestaña por mes con una cuadrícula de empleados por día.
- Envía un correo con el resumen, las personas que llegaron tarde y las que no registraron entrada.

---

## 2. Requisitos

| Requisito | Detalle |
|---|---|
| Cuenta de Google | La cuenta que ejecuta el script debe ser miembro del espacio de Chat y tener acceso de edición al Google Sheet. |
| Google Sheet | Un archivo de Google Sheets. El script lo abre por su ID, no por estar vinculado a él. |
| Servicio avanzado | **Google Chat API v1**, con identificador `Chat`. |
| Navegador | Acceso a [script.google.com](https://script.google.com). |

---

## 3. Valores que se editan en el código

Al inicio de `v9.js` está el objeto `ATTENDANCE`. Hay tres valores propios de tu entorno:

| Campo | Para qué sirve |
|---|---|
| `spreadsheetId` | ID del Google Sheet donde se guarda todo. Es la parte de la URL entre `/d/` y `/edit`. |
| `defaultSpace` | Espacio de Chat (`spaces/XXXX`). Solo se usa para llenar `Config!B4` la primera vez. |
| `timezone` | Zona horaria con la que se calcula el día y las horas. Hoy es `America/Mexico_City`. |

Además hay listas que se pueden ajustar: `builtInPhrases` (frases que siempre cuentan como asistencia), `fillerWords` (palabras que pueden ir antes del saludo), `excludedNextWords` (palabras que anulan el saludo, como "tardes") y `colors`.

---

## 4. Primera ejecución paso a paso

1. **Abre el proyecto de Apps Script** y reemplaza todo el contenido de `Code.gs` por el de `v9.js`.
2. **Activa el servicio de Chat:** en el panel izquierdo, junto a *Servicios*, pulsa `+`, elige **Google Chat API** y agrégalo con versión **v1** e identificador `Chat`.
3. **Revisa `spreadsheetId` y `timezone`** en el objeto `ATTENDANCE`.
4. **Ejecuta `testConnection`.** Elige la función en el menú desplegable de la barra superior y pulsa **Ejecutar**. La primera vez Google pide autorización: acepta los permisos. Debe aparecer en el registro "CONNECTION SUCCESSFUL" con las pestañas que ya existen.
5. **Ejecuta `setupAttendanceTracker`.** Crea las pestañas que falten y deja todo configurado (ver sección 5).
6. **Completa `Config`** en el Google Sheet:
   - `B2` (fecha): **déjala vacía** para que se use el día de hoy. En una hoja nueva el setup escribe una fecha de ejemplo (`2025-04-29`); bórrala.
   - `B4`: confirma el espacio de Chat.
   - `B5`: hora límite de entrada, por ejemplo `09:15`.
   - `B6`: correos para el reporte, o déjala vacía si no quieres correo.
7. **Ejecuta `importEmployeesFromChat`.** Llena la pestaña `Employees` con los miembros del espacio.
8. **Revisa `Employees`:** si alguien no debe contar, cambia su columna `Active` a `FALSE`.
9. **Ejecuta `runAttendance`.** Procesa el día y actualiza todas las pestañas.

Si algo falla, abre **Registro de ejecución** para ver el mensaje (sección 11).

---

## 5. Funciones que se ejecutan a mano

| Función | Qué hace | Por qué se ejecuta | Cuándo |
|---|---|---|---|
| `testConnection` | Abre el Google Sheet y muestra en el registro sus pestañas y cuántas filas tiene cada una. | Comprobar que el ID es correcto y que hay permisos. | Una vez al inicio, o al diagnosticar un problema de acceso. |
| `setupAttendanceTracker` | Crea o completa las pestañas `Config`, `Employees`, `Attendance` y `Vacations`, sus encabezados, anchos, listas desplegables, filtro y las 12 pestañas mensuales del año en curso. No borra datos existentes. | Dejar la hoja lista para trabajar. | Una vez al inicio. Es seguro repetirla. Conviene volver a correrla al empezar un año nuevo (aunque `runAttendance` también crea la pestaña del mes si falta). |
| `importEmployeesFromChat` | Agrega a `Employees` los miembros humanos del espacio que aún no están. No modifica ni elimina a los existentes. | Mantener la lista de empleados al día. | Al inicio y cada vez que entre alguien nuevo al espacio. |
| `runAttendance` | Procesa un día completo (ver sección 6). | Es la función principal: genera el registro de asistencia. | Una vez al día, después de la hora límite o al terminar la jornada. |
| `applyAttendanceFormatting` | Vuelve a pintar las filas de `Attendance` según su estado. | Solo si los colores se desordenaron. | Casi nunca; `setupAttendanceTracker` y `runAttendance` ya la llaman. |

El resto de las funciones terminan en `_` y son internas: no hay que ejecutarlas.

---

## 6. Qué hace `runAttendance` paso a paso

1. Impide que corran dos ejecuciones a la vez (espera hasta 10 segundos).
2. Lee `Config`: fecha, frases, espacio, hora límite y correos. Valida todo **antes** de escribir datos.
3. Calcula el rango del día en la zona `ATTENDANCE.timezone`.
4. Descarga los mensajes del espacio en ese rango. Ignora bots y mensajes de otro día.
5. Para cada mensaje decide si es asistencia (sección 8) y conserva el **más temprano** de cada persona.
6. Lee `Employees` y `Vacations`.
7. Genera una fila por empleado activo, más las personas que escribieron sin estar en la lista.
8. Sustituye en `Attendance` las filas de esa fecha y conserva el resto del historial, ordenado de la fecha más reciente a la más antigua.
9. Aplica colores y filtro a `Attendance`.
10. Marca el día en la pestaña del mes.
11. Envía el correo, si hay destinatarios y algo que reportar.

### Fecha que se procesa

- `Config!B2` **vacía**: se procesa el día de hoy.
- `Config!B2` con una fecha `YYYY-MM-DD`: se procesa esa fecha.

> **Importante:** si escribes una fecha en `B2` para reprocesar un día pasado, `runAttendance` seguirá usando esa fecha hasta que **borres la celda**. Después de corregir un día, vacíala.

### Estados que se registran

| Estado | Significado | Color |
|---|---|---|
| `PRESENT` | Registró entrada antes o a la hora límite. | Verde |
| `PRESENT` (tarde) | Registró entrada después de la hora límite. Es tarde si la hora es mayor a `Late After`, comparada al segundo (09:15:08 con corte 09:15 cuenta como tarde). | Rojo |
| `NO CHECK-IN` | Empleado activo sin registro válido. | Ámbar |
| `PRESENT - NOT ON ROSTER` | Registró entrada pero no está en `Employees`. | Azul |
| `ON VACATION` | Tiene vacaciones aprobadas que cubren la fecha. | Índigo |

---

## 7. Pestañas del Google Sheet

### `Config`

| Celda | Contenido | Formato |
|---|---|---|
| `B2` | Fecha a procesar. Vacía = hoy. | `YYYY-MM-DD` |
| `B3` | Frases adicionales de asistencia, separadas por coma. | Texto |
| `B4` | Espacio de Chat. | `spaces/XXXX` |
| `B5` | Hora límite. | `HH:mm`, sin segundos |
| `B6` | Correos del reporte, separados por coma o punto y coma. Vacía = no se envía. | Texto |

Las celdas `B2:B6` se fuerzan a formato texto para que Sheets no convierta la fecha ni la hora.

### `Employees`

Columnas: `Google Chat User ID`, `Employee Name`, `Active`. Los valores que cuentan como activo son vacío, `TRUE`, `YES`, `SI`, `SÍ` y `1`. No puede haber IDs repetidos: si los hay, `runAttendance` se detiene.

### `Attendance`

Columnas: `Date`, `Google Chat User ID`, `Employee`, `Status`, `First Check-in`, `Message ID`. Tiene un filtro en el encabezado: usa la flecha de `Date` para ver un día y la de `Employee` para ver a una persona; se pueden combinar. Cada ejecución recrea el filtro para que cubra todas las filas y conserva los criterios que tenías activos.

### `Vacations`

Columnas: `Employee`, `Start Date`, `End Date`, `Status`.

- `Status` es una lista cerrada: `APPROVED`, `PENDING`, `REJECTED`.
- Solo las filas `APPROVED` cuyo rango incluye la fecha procesada cuentan. Las demás se ignoran y el empleado sigue el flujo normal.
- Las fechas se escriben `YYYY-MM-DD` (inicio y fin incluidos; para un solo día, la misma fecha en ambas).
- El empleado se relaciona **por nombre** con `Employees` (sin importar mayúsculas ni acentos). La columna `Employee` tiene lista desplegable con los nombres.
- Una fila `APPROVED` con fechas inválidas o sin nombre **detiene** la ejecución e indica la fila. Si el nombre no existe en `Employees`, solo se registra una advertencia.
- Si alguien de vacaciones también escribe su entrada, queda como `ON VACATION` y se conserva su hora como referencia.

### Pestañas mensuales

`setupAttendanceTracker` crea una por mes del año en curso (por ejemplo `Noviembre 2026`). Si procesas una fecha de otro año, esa pestaña se crea al momento.

- Una fila por empleado (columna A) y una columna por **día hábil** (lunes a viernes), con encabezado `DD/MM/YYYY`.
- `X` verde: entrada a tiempo. `X` roja: entrada tarde. `V`: vacaciones. Celda vacía: sin registro.
- Si procesas un sábado o domingo, no se escribe nada en la pestaña del mes (queda una advertencia en el registro). `Attendance` sí se actualiza.
- Los empleados nuevos se agregan al final en la siguiente corrida. Quienes no están en `Employees` no aparecen aquí.

---

## 8. Cómo se detecta un mensaje de asistencia

Un mensaje cuenta cuando **empieza** con una frase válida. Las frases válidas son las de `Config!B3` más las integradas (`buen dia`, `buenos dias`, `buenas dias`, `buenas`, `good morning`, `good day`, `presente`, `entrada`, `reporto entrada`, `reporto mi entrada`).

| Regla | Ejemplo que cuenta |
|---|---|
| Ignora mayúsculas, acentos, signos y emojis | `¡Buenos días!`, `Buen día 🤟` |
| Permite palabras antes del saludo | `Muy buenos dias`, `Hola, buen día` |
| Permite texto después | `Buen día, reporto mi entrada.`, `Buenos días !! Excelente jueves !!` |
| Tolera errores pequeños | `Gooday`, `Buenooos días` |

No cuenta:

- Si la frase aparece en medio del texto: `Que tengan un excelente día`.
- `Buenas tardes` y `Buenas noches`.
- Respuestas como `Enterado` o `Gracias`.

Si un empleado quedó como `NO CHECK-IN` pero escribió algo ese día, el registro de ejecución muestra una línea `NO CHECK-IN but wrote messages | nombre | texto`. Úsala para descubrir frases que faltan y agrégalas en `Config!B3`.

Limitación: cualquier mensaje que empiece con un saludo cuenta como asistencia, aunque sea una respuesta a otra conversación.

---

## 9. Reporte por correo

- Se envía al final de `runAttendance` a los correos de `Config!B6`.
- Solo se envía si hay al menos una persona tarde o sin registro. Si todos llegaron a tiempo, no se manda nada.
- Contiene: fecha, hora límite, totales (a tiempo, tarde, vacaciones, sin registro), tabla de llegadas tarde con hora y retraso, lista de quienes no registraron entrada y un botón para abrir la hoja.
- Las personas de vacaciones no aparecen como "sin registro".
- **Se envía en cada ejecución.** Si corres el mismo día dos veces, llegan dos correos.
- Un correo mal escrito en `B6` detiene la ejecución antes de escribir datos.
- La primera vez que corra `runAttendance` con correos configurados, Google pedirá permiso para enviar correo.
- El correo sale desde la cuenta que ejecuta el script.

---

## 10. Cada cuánto ejecutar

| Cuándo | Qué hacer |
|---|---|
| Una sola vez | `testConnection`, `setupAttendanceTracker` e `importEmployeesFromChat`. |
| Cada día hábil | `runAttendance`, **después** de la hora límite o al cerrar la jornada. Si lo corres antes de que todos escriban, los que faltan aparecen como `NO CHECK-IN` y el correo los incluye. |
| Cuando entren personas nuevas al espacio | `importEmployeesFromChat`. Después, si hace falta, `runAttendance` de nuevo. |
| Cuando alguien salga del espacio | Cambiar a mano su columna `Active` a `FALSE` en `Employees`. El script no lo hace solo. |
| Para corregir o recalcular un día pasado | Escribir la fecha en `Config!B2`, ejecutar `runAttendance` y **vaciar `B2`** al terminar. |
| Al empezar un año nuevo | Opcional: `setupAttendanceTracker` para crear de golpe las 12 pestañas. |

Re-ejecutar el mismo día es seguro: solo se reemplazan las filas de esa fecha y el resto del historial se conserva. Pero **las ediciones manuales a esas filas, y a la celda del empleado en la pestaña del mes, se sobrescriben**.

### Ejecución automática (opcional, no incluida)

El script no crea ningún activador. Para que corra solo cada día: en Apps Script abre **Activadores** → **Agregar activador** → función `runAttendance`, tipo *Basado en tiempo* → *Temporizador diario*, a una hora posterior al corte. Con `Config!B2` vacía procesará el día de hoy. Ten en cuenta que Apps Script ejecuta el activador en cualquier momento dentro de la hora elegida y que también correrá los fines de semana y días festivos, donde todos saldrían como `NO CHECK-IN`.

---

## 11. Errores frecuentes

| Mensaje | Causa y solución |
|---|---|
| `Cannot access spreadsheet: ...` | `spreadsheetId` incorrecto o la cuenta no tiene acceso al archivo. |
| `Missing required tabs. Run setupAttendanceTracker() first.` | Falta alguna pestaña, incluida `Vacations`. Ejecuta `setupAttendanceTracker`. |
| `Invalid Config!B4 space ID.` | `B4` debe tener el formato `spaces/XXXX`. |
| `Config!B2 must use YYYY-MM-DD or be empty.` / `Invalid attendance date.` | Fecha mal escrita o inexistente en `B2`. |
| `Config!B5 must contain a time such as 09:15.` | `B5` debe ser `HH:mm`, sin segundos. |
| `Config!B6 has an invalid email: ...` | Algún correo de `B6` no es válido. |
| `Duplicate employee ID: ...` | Hay un ID repetido en `Employees`. |
| `Vacations row N: ...` | Una fila `APPROVED` tiene datos inválidos (nombre vacío, fechas con otro formato o inicio posterior al fin). |
| `Another attendance execution is running.` | Hay otra ejecución en curso. Espera y vuelve a intentarlo. |
| Error de la API de Chat o de autorización | Revisa que el servicio `Chat` v1 esté agregado, que aceptaste los permisos y que tu cuenta pertenece al espacio. |

---

## 12. Limitaciones conocidas

- No distingue fines de semana ni días festivos: si ejecutas `runAttendance` esos días, todos salen sin registro.
- `Vacations` y las pestañas mensuales relacionan al empleado por nombre. Si cambia un nombre, aparecerá una fila nueva y la anterior quedará sin actualizar.
- Los empleados inactivos no generan fila, aunque escriban su entrada.
- Quienes escriben sin estar en `Employees` aparecen en `Attendance` (en azul), pero no en las pestañas mensuales.
- Las listas desplegables de `Employees`/`Vacations` cubren hasta la fila 1000.
- `Attendance` se reescribe en cada ejecución (se borra el contenido y se vuelve a escribir). Si necesitas recuperar datos, usa *Archivo → Historial de versiones* en Google Sheets.
- Los IDs y el espacio de Chat están escritos en el código. Para usarlo con otro archivo o espacio hay que editarlos.
- El corte de hora compara al segundo, sin tolerancia.

---

## 13. Estado de las pruebas

`v9.js` pasa la verificación de sintaxis de Node (`node --check`) y su lógica se probó con simulaciones locales (detección de mensajes, vacaciones, pestañas mensuales, filtro y generación del correo). **No se ha ejecutado todavía contra Google Chat ni en Apps Script real.** Antes de usarlo en producción, pruébalo en una copia del Google Sheet con una fecha pasada que conozcas y compara el resultado con el chat.

---

## 14. Historial de versiones

| Archivo | Cambios |
|---|---|
| `v1.js` | Versión original. |
| `v2.js` | Rango del día con la zona horaria (sin `-06:00` fijo), formato texto forzado en `Config` y colores leídos de `ATTENDANCE.colors`. |
| `v3.js` | Pestaña `Vacations` y estado `ON VACATION`. |
| `v4.js` | `Config!B2` vacía procesa el día de hoy. |
| `v5.js` | Filtro en `Attendance` por fecha y empleado. |
| `v6.js` | Una pestaña por mes con `X` verde/roja y `V`. |
| `v7.js` | Detección flexible de mensajes y reporte de mensajes ignorados. |
| `v8.js` | Reporte por correo (`Config!B6`). |
| `v9.js` | Correo con diseño mejorado: resumen, tablas, retraso y botón. |

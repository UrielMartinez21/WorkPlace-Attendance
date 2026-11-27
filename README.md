# Control de asistencia con Google Chat y Google Sheets (`v10.js`)

Script de Google Apps Script que lee los mensajes de saludo de un espacio de Google Chat, decide quién registró su entrada y quién no, y deja el resultado en una hoja de Google Sheets. También envía un correo con quienes llegaron tarde o no reportaron, publica un aviso diario en el grupo de Chat y mantiene un dashboard.

Este documento describe la versión **`v10.js`**, la última de la carpeta. Se usa junto con `appsscript.json`, que declara el servicio avanzado y los permisos.

---

## 1. Qué hace

- Importa los miembros del espacio de Chat a una lista de empleados, cada uno con un país (México o Colombia).
- Busca los mensajes de un día y detecta cuáles son un registro de asistencia (por ejemplo "Buen día", "Buenos días, reporto mi entrada", "Presente").
- Conserva el **primer** mensaje válido de cada persona en el día.
- Marca a cada empleado como `PRESENT` (a tiempo o tarde), `NO CHECK-IN`, `ON VACATION` o `HOLIDAY`.
- Importa los días festivos de México y Colombia, y no cuenta como falta a quien tiene festivo en su país.
- Guarda el historial en la pestaña `Attendance` (con filtro y colores).
- Rellena una pestaña por mes con una cuadrícula de empleados por día hábil.
- Mantiene una pestaña `Dashboard` con porcentajes, tendencias, rankings y filtros por mes, país y empleado.
- Envía un correo con el resumen, las personas que llegaron tarde y las que no registraron entrada.
- Publica en el grupo de Chat un aviso diario con quién atiende WhatsApp, quién está de vacaciones y quién lo estará en 3 días hábiles (sección 15).

---

## 2. Requisitos

| Requisito | Detalle |
|---|---|
| Cuenta de Google | La cuenta que ejecuta el script debe ser miembro del espacio de Chat y tener acceso de edición al Google Sheet. |
| Google Sheet | Un archivo de Google Sheets. El script lo abre por su ID, no por estar vinculado a él. |
| Servicio avanzado | **Google Chat API v1**, con identificador `Chat`. |
| Manifiesto | `appsscript.json` con los permisos (scopes) de Chat, Sheets, correo, calendario y `script.external_request`. |
| Webhook de Chat | Para el aviso diario. Ver sección 15. |
| Calendario de festivos | `importHolidays` usa los calendarios públicos de festivos de Google. Pide permiso la primera vez. |
| Navegador | Acceso a [script.google.com](https://script.google.com). |

---

## 3. Valores que se editan en el código

Al inicio de `v10.js` está el objeto `ATTENDANCE`. Hay tres valores propios de tu entorno:

| Campo | Para qué sirve |
|---|---|
| `spreadsheetId` | ID del Google Sheet donde se guarda todo. Es la parte de la URL entre `/d/` y `/edit`. |
| `defaultSpace` | Espacio de Chat (`spaces/XXXX`). Solo se usa para llenar `Config!B4` la primera vez. |
| `timezone` | Zona horaria con la que se calcula el día y las horas. Hoy es `America/Mexico_City`. |

Otros valores que se pueden ajustar:

| Campo | Para qué sirve |
|---|---|
| `countries` y `defaultCountry` | Países admitidos en `Employees!D` (`MX`, `CO`) y el que se asigna a los empleados nuevos (`MX`). |
| `holidayCalendars` | IDs de los calendarios públicos de festivos que usa `importHolidays`. |
| `holidayExcludePattern` | Eventos de esos calendarios que no son días no laborables (Día de la Madre, etc.) y se omiten. |
| `builtInPhrases` | Frases que siempre cuentan como asistencia. |
| `fillerWords` | Palabras que pueden ir antes del saludo. |
| `excludedNextWords` | Palabras que anulan el saludo, como "tardes". |
| `colors` | Colores de estados, encabezados y pestañas mensuales. |
| `VACATION_NOTICE_DAYS_AHEAD` | Días hábiles de anticipación del aviso de vacaciones (3). Está en la sección `AVISO DE VACACIONES EN GOOGLE CHAT`. |

La URL del webhook de Chat **no** se escribe en el código. Se guarda como propiedad del script (sección 15).

---

## 4. Primera ejecución paso a paso

1. **Abre el proyecto de Apps Script.** Reemplaza todo el contenido de `Code.gs` por el de `v10.js`. Reemplaza `appsscript.json` por el de la carpeta (si no lo ves, actívalo en *Configuración del proyecto → Mostrar el archivo de manifiesto*).
2. **Activa el servicio de Chat:** en el panel izquierdo, junto a *Servicios*, pulsa `+`, elige **Google Chat API** y agrégalo con versión **v1** e identificador `Chat`.
3. **Revisa `spreadsheetId` y `timezone`** en el objeto `ATTENDANCE`.
4. **Ejecuta `testConnection`.** Elige la función en el menú desplegable de la barra superior y pulsa **Ejecutar**. La primera vez Google pide autorización: acepta los permisos. Debe aparecer en el registro "CONNECTION SUCCESSFUL" con las pestañas que ya existen.
5. **Ejecuta `setupAttendanceTracker`.** Crea las pestañas que falten y deja todo configurado (ver sección 5).
6. **Completa `Config`** en el Google Sheet:
   - `B4`: confirma el espacio de Chat.
   - `B5`: hora límite de entrada, por ejemplo `09:15`.
   - `B6`: correos para el reporte, o déjala vacía si no quieres correo.
   - `B2` solo se usa para procesar una fecha pasada con `runAttendanceForDate`. En una hoja nueva el setup escribe una fecha de ejemplo (`2025-04-29`). La ejecución diaria la ignora.
7. **Ejecuta `importEmployeesFromChat`.** Llena la pestaña `Employees` con los miembros del espacio. Los nuevos quedan con `Country = MX`.
8. **Revisa `Employees`:** si alguien no debe contar, cambia su columna `Active` a `FALSE`. Cambia `Country` a `CO` a quienes trabajan en Colombia.
9. **Ejecuta `importHolidays`.** Llena la pestaña `Holidays` con los festivos del año en curso y el siguiente. Revísala.
10. **Configura el aviso de Chat** (sección 15): crea el webhook, guarda `CHAT_WEBHOOK_URL` y ejecuta `postDailyNotice` para autorizar y probar.
11. **Prueba `runAttendanceForDate`** con una fecha pasada en `Config!B2` y compara con el chat. Después vacía `B2`.
12. **Crea los dos activadores diarios**: `runDailyAttendance` y `postDailyNotice` (sección 10).

Si algo falla, abre **Registro de ejecución** para ver el mensaje (sección 11).

---

## 5. Funciones que se ejecutan a mano

| Función | Qué hace | Cuándo |
|---|---|---|
| `testConnection` | Abre el Google Sheet y muestra en el registro sus pestañas y cuántas filas tiene cada una. | Una vez al inicio, o al diagnosticar un problema de acceso. |
| `setupAttendanceTracker` | Crea o completa las pestañas `Config`, `Employees`, `Attendance`, `Vacations`, `Holidays` y `WhatsApp`, con encabezados, anchos, listas desplegables y filtro. Crea las 12 pestañas mensuales del año en curso y refresca el dashboard. No borra datos existentes, salvo la columna `DIA` de versiones anteriores de `WhatsApp`. | Una vez al inicio. Es seguro repetirla. Conviene volver a correrla al empezar un año nuevo. |
| `importEmployeesFromChat` | Agrega a `Employees` los miembros humanos del espacio que aún no están, con `Country = MX`. No modifica ni elimina a los existentes. | Al inicio y cada vez que entre alguien nuevo al espacio. |
| `importHolidays` | Agrega a `Holidays` los festivos de México y Colombia del año en curso y el siguiente. Solo agrega fechas que faltan. | Una vez al inicio y al empezar cada año. |
| `runAttendanceForDate` | Procesa la fecha escrita en `Config!B2` (obligatoria). Acepta cualquier fecha, fines de semana incluidos. No publica en Chat. | Para corregir o recalcular un día pasado. |
| `runDailyAttendance` | Procesa siempre el día de hoy e ignora `Config!B2`. No hace nada en sábado ni domingo. No publica en Chat. | Es la función del activador de asistencia. |
| `postDailyNotice` | Publica el aviso de hoy en el grupo de Chat. Es independiente de la asistencia. No hace nada en sábado ni domingo. | Es la función del activador del aviso. También sirve para autorizar el permiso y probar el mensaje. |
| `refreshDashboard` | Reconstruye los datos y fórmulas del dashboard. | Casi nunca: ambas ejecuciones de asistencia ya lo hacen. Los filtros funcionan solos. |
| `applyAttendanceFormatting` | Vuelve a pintar las filas de `Attendance` según su estado. | Solo si los colores se desordenaron. |

El resto de las funciones terminan en `_` y son internas: no hay que ejecutarlas.

---

## 6. Qué hace una ejecución de asistencia

`runDailyAttendance` y `runAttendanceForDate` siguen los mismos pasos. Se diferencian en la fecha que procesan. Ninguna publica en Chat: el aviso lo envía `postDailyNotice`.

1. Impide que corran dos ejecuciones a la vez (espera hasta 10 segundos).
2. Lee `Config`: fecha, frases, espacio, hora límite y correos. Valida todo **antes** de escribir datos.
3. En modo diario, si hoy es sábado o domingo, termina sin escribir ni enviar nada.
4. Calcula el rango del día en la zona `ATTENDANCE.timezone`.
5. Descarga los mensajes del espacio en ese rango. Ignora bots y mensajes de otro día.
6. Para cada mensaje decide si es asistencia (sección 8) y conserva el **más temprano** de cada persona.
7. Lee `Employees`, `Vacations` y `Holidays`.
8. Genera una fila por empleado activo, más las personas que escribieron sin estar en la lista.
9. Sustituye en `Attendance` las filas de esa fecha y conserva el resto del historial, ordenado de la fecha más reciente a la más antigua.
10. Aplica colores y filtro a `Attendance`.
11. Marca el día en la pestaña del mes.
12. Envía el correo, si hay destinatarios y algo que reportar.
13. En modo diario, publica el aviso en el grupo de Chat.
14. Refresca el dashboard. Si falla, solo deja una advertencia: la asistencia ya está guardada.

### Fecha que se procesa

| Función | Fecha |
|---|---|
| `runDailyAttendance` | Hoy, en la zona `ATTENDANCE.timezone`. Ignora `Config!B2`. |
| `runAttendanceForDate` | La de `Config!B2`, con formato `YYYY-MM-DD`. Si está vacía, la ejecución se detiene con un error. |

### Estados que se registran

| Estado | Significado | Color |
|---|---|---|
| `PRESENT` | Registró entrada antes o a la hora límite. | Verde |
| `PRESENT` (tarde) | Registró entrada después de la hora límite. Es tarde si la hora es mayor a `Late After`, comparada al segundo (09:15:08 con corte 09:15 cuenta como tarde). | Rojo |
| `NO CHECK-IN` | Empleado activo sin registro válido. | Ámbar |
| `PRESENT - NOT ON ROSTER` | Registró entrada pero no está en `Employees`. | Azul |
| `ON VACATION` | Tiene vacaciones aprobadas que cubren la fecha. | Índigo |
| `HOLIDAY` | La fecha es festivo en el país del empleado. | Gris |
| `EXEMPT` | El empleado está en `Config!B7` y no debe reportar asistencia. | Verde azulado |

`EXEMPT` tiene prioridad sobre `HOLIDAY`, y `HOLIDAY` sobre `ON VACATION`. En todos los casos, si la persona escribió su entrada, se conserva su hora como referencia.

---

## 7. Pestañas del Google Sheet

### `Config`

| Celda | Contenido | Formato |
|---|---|---|
| `B2` | Fecha a procesar con `runAttendanceForDate`. | `YYYY-MM-DD` |
| `B3` | Frases adicionales de asistencia, separadas por coma. | Texto |
| `B4` | Espacio de Chat que se lee. | `spaces/XXXX` |
| `B5` | Hora límite. | `HH:mm`, sin segundos |
| `B6` | Correos del reporte, separados por coma o punto y coma. Vacía = no se envía. | Texto |
| `B7` | Empleados exentos de reportar asistencia (`Exempt Employees`). Vacía = nadie está exento. | Nombres separados por coma, punto y coma o salto de línea |

Las celdas `B2:B7` se fuerzan a formato texto para que Sheets no convierta la fecha ni la hora.

**Lista de áreas (`C2:C50`).** La columna `C` (encabezado `Areas`) guarda las áreas disponibles, una por celda. Alimenta la lista desplegable de `Employees!E`. Escríbelas a mano; el script solo crea el encabezado.

**Empleados exentos (`B7`).** Escribe los nombres tal como aparecen en `Employees!B` (no importan mayúsculas ni acentos), por ejemplo `González Rodríguez Joel; Carrillo Vélez Betzy`. Estas personas quedan con estado `EXEMPT`:

- No cuentan como `NO CHECK-IN` ni como tarde, aunque no escriban nada. Si escriben su entrada, se conserva su hora solo como referencia.
- `EXEMPT` tiene prioridad sobre `HOLIDAY` y `ON VACATION`.
- No aparecen en el correo ni en el dashboard, y en la pestaña del mes llevan una `E` verde azulado.
- Se siguen relacionando por nombre: si cambia el nombre en `Employees`, actualiza también `B7`. Un nombre de `B7` que no coincide con ningún empleado activo deja una advertencia en el registro.
- Para quitar la exención, borra el nombre de `B7`. Las filas de días ya procesados conservan `EXEMPT` hasta que los reprocesas con `runAttendanceForDate`.

### `Employees`

Columnas: `Google Chat User ID`, `Employee Name`, `Active`, `Country`, `Area`.

- Los valores que cuentan como activo son vacío, `TRUE`, `YES`, `SI`, `SÍ` y `1`.
- `Country` es `MX` o `CO` (lista desplegable). Vacío equivale a `MX`. Define qué festivos aplican al empleado.
- `Area` es una lista desplegable con las áreas de `Config!C2:C50`. Solo se usa en el correo. Es opcional y la validación solo advierte. `importEmployeesFromChat` la deja vacía en los empleados nuevos.
- No puede haber IDs repetidos: si los hay, la ejecución se detiene.

### `Attendance`

Columnas: `Date`, `Google Chat User ID`, `Employee`, `Status`, `First Check-in`, `Message ID`. Tiene un filtro en el encabezado: usa la flecha de `Date` para ver un día y la de `Employee` para ver a una persona; se pueden combinar. Cada ejecución recrea el filtro para que cubra todas las filas y conserva los criterios que tenías activos.

### `Vacations`

Columnas: `Employee`, `Start Date`, `End Date`, `Status`.

- `Status` es una lista cerrada: `APPROVED`, `PENDING`, `REJECTED`.
- Solo las filas `APPROVED` cuyo rango incluye la fecha procesada cuentan. Las demás se ignoran y el empleado sigue el flujo normal.
- `Start Date` y `End Date` tienen selector de calendario (doble clic en la celda) y se muestran como `YYYY-MM-DD`. Inicio y fin están incluidos; para un solo día, la misma fecha en ambas. La validación solo advierte: también se leen fechas escritas o pegadas como texto `YYYY-MM-DD`.
- El empleado se relaciona **por nombre** con `Employees` (sin importar mayúsculas ni acentos). La columna `Employee` tiene lista desplegable con los nombres.
- Una fila `APPROVED` con fechas inválidas o sin nombre **detiene** la ejecución e indica la fila. Si el nombre no existe en `Employees`, solo se registra una advertencia.
- Esta pestaña también alimenta el aviso diario de Chat (sección 15).

### `Holidays`

Columnas: `Date`, `Country`, `Name`.

- `Date` va como `YYYY-MM-DD`. `Country` es `MX`, `CO` o `ALL` (aplica a todos los países). `Name` es el nombre del festivo.
- `importHolidays` la llena desde los calendarios públicos de Google. Después puedes editarla a mano para añadir o quitar festivos.
- Una fila con fecha o país inválidos **detiene** la ejecución e indica la fila.
- Si no hay festivos cargados para el año procesado, la ejecución continúa y deja una advertencia.

### `WhatsApp`

Columnas: `NOMBRE` y `FECHA`. Indica quién atiende los mensajes de WhatsApp cada día. Ver sección 15.

### Pestañas mensuales

`setupAttendanceTracker` crea una por mes del año en curso (por ejemplo `Noviembre 2026`). Si procesas una fecha de otro año, esa pestaña se crea al momento.

- Una fila por empleado (columna A) y una columna por **día hábil** (lunes a viernes), con encabezado `DD/MM/YYYY`.
- `X` verde: entrada a tiempo. `X` roja: entrada tarde. `V`: vacaciones. `F` gris: festivo. `E` verde azulado: exento. Celda vacía: sin registro.
- Si procesas un sábado o domingo con `runAttendanceForDate`, no se escribe nada en la pestaña del mes (queda una advertencia en el registro). `Attendance` sí se actualiza.
- Los empleados nuevos se agregan al final en la siguiente corrida. Quienes no están en `Employees` no aparecen aquí.

### `Dashboard` y `DashData`

`Dashboard` muestra tarjetas, cinco gráficas, dos rankings y una tabla por empleado. `DashData` guarda una fila por empleado y día, y alimenta las fórmulas. No hay que editar `DashData`.

- Los filtros están en `Dashboard!B3:B5` (mes, país y empleado). Funcionan con fórmulas: al cambiar un filtro se actualiza al instante, sin ejecutar nada.
- El script escribe los datos y las fórmulas en cada ejecución de asistencia (o con `refreshDashboard`).
- Reglas de cálculo:
  - Sábados y domingos no cuentan.
  - Vacaciones y festivos no cuentan como falta ni bajan los porcentajes: se muestran aparte.
  - `% asistencia` = (a tiempo + tarde) / (a tiempo + tarde + sin registro).
  - `% puntualidad` = a tiempo / (a tiempo + tarde).
  - "Tarde" usa la hora límite de `Config!B5` vigente en la última ejecución.
  - El país sale de `Employees!D` actual, no el de la fecha.
  - `PRESENT - NOT ON ROSTER` y `EXEMPT` no se incluyen.

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

- Se envía al final de cada ejecución de asistencia (diaria o por fecha) a los correos de `Config!B6`.
- Solo se envía si hay al menos una persona tarde o sin registro. Si todos llegaron a tiempo, no se manda nada.
- Contiene: fecha, hora límite, totales (a tiempo, tarde, vacaciones, festivo y sin registro), tabla de llegadas tarde con empleado, área, país, hora y retraso, tabla de quienes no registraron entrada con empleado, área y país, tabla de quienes están de vacaciones con empleado, área y país (solo si hay alguien), y un botón para abrir la hoja.
- El área sale de `Employees!E` y el país de `Employees!D`. Ambos aparecen en las tablas de llegadas tarde, sin registro y vacaciones, como etiquetas de color. Quien no tenga área asignada aparece como "Sin área" en gris. La tabla sigue ordenada por hora de llegada, sin agrupar por área.
- Colores: cada área toma un color según su posición en `Config!C2:C50` (hay 6 colores y la séptima área repite el primero), así que conviene no reordenar esa lista. El país tiene color fijo (México azul cielo, Colombia amarillo). Los colores se editan en `ATTENDANCE.colors` (`areaPalette`, `countryColors`, `neutralBadge`).
- Si hay festivo, el correo muestra una tarjeta de "Festivo" y el nombre del festivo por país.
- Las personas de vacaciones o con festivo no aparecen como "sin registro". Las de vacaciones se listan aparte, en la sección "De vacaciones", que solo aparece si hay alguien. No se muestran sus fechas de regreso.
- La sección de vacaciones no hace que se envíe el correo: sigue enviándose solo si hay alguien tarde o sin registro.
- **Se envía en cada ejecución.** Si corres el mismo día dos veces, llegan dos correos.
- Un correo mal escrito en `B6` detiene la ejecución antes de escribir datos.
- La primera vez que corra una ejecución con correos configurados, Google pedirá permiso para enviar correo.
- El correo sale desde la cuenta que ejecuta el script.

---

## 10. Cada cuánto ejecutar

| Cuándo | Qué hacer |
|---|---|
| Una sola vez | `testConnection`, `setupAttendanceTracker`, `importEmployeesFromChat`, `importHolidays` y configurar el webhook (sección 15). |
| Cada día hábil | `runDailyAttendance`, **después** de la hora límite o al cerrar la jornada. Si corre antes de que todos escriban, los que faltan aparecen como `NO CHECK-IN` y el correo los incluye. Conviene dejarla en un activador. |
| Cada día hábil, por la mañana | `postDailyNotice`, en su propio activador. Publica el aviso de WhatsApp y vacaciones. |
| Cuando entren personas nuevas al espacio | `importEmployeesFromChat`. Revisa su `Country`. Después, si hace falta, ejecuta de nuevo la asistencia. |
| Cuando alguien salga del espacio | Cambiar a mano su columna `Active` a `FALSE` en `Employees`. El script no lo hace solo. |
| Cada año nuevo | `importHolidays` (también carga el año siguiente) y, opcional, `setupAttendanceTracker` para crear de golpe las 12 pestañas. |
| Para corregir o recalcular un día pasado | Escribir la fecha en `Config!B2` y ejecutar `runAttendanceForDate`. No publica en Chat. |

Re-ejecutar el mismo día es seguro: solo se reemplazan las filas de esa fecha y el resto del historial se conserva. Pero **las ediciones manuales a esas filas, y a la celda del empleado en la pestaña del mes, se sobrescriben**.

### Activadores diarios

Son dos activadores independientes. En Apps Script abre **Activadores** → **Agregar activador**, elige el tipo *Basado en tiempo* → *Temporizador diario* y crea:

| Función | Hora sugerida |
|---|---|
| `runDailyAttendance` | Posterior a la hora límite de entrada. |
| `postDailyNotice` | La que prefieras para el aviso, por ejemplo temprano por la mañana. No depende de la asistencia. |

Apps Script ejecuta cada activador en cualquier momento dentro de la hora elegida. Ambas funciones no hacen nada en fin de semana. Los festivos de asistencia se resuelven por la pestaña `Holidays` y el país de cada empleado. El aviso de Chat no considera festivos.

---

## 11. Errores frecuentes

| Mensaje | Causa y solución |
|---|---|
| `Cannot access spreadsheet: ...` | `spreadsheetId` incorrecto o la cuenta no tiene acceso al archivo. |
| `Missing required tabs. Run setupAttendanceTracker() first.` | Falta alguna pestaña (`Config`, `Employees`, `Attendance`, `Vacations` o `Holidays`). Ejecuta `setupAttendanceTracker`. |
| `Invalid Config!B4 space ID.` | `B4` debe tener el formato `spaces/XXXX`. |
| `Type the date to process in Config!B2 (YYYY-MM-DD).` | Ejecutaste `runAttendanceForDate` con `B2` vacía. |
| `Config!B2 must use YYYY-MM-DD.` / `Invalid attendance date.` | Fecha mal escrita o inexistente en `B2`. |
| `Config!B5 must contain a time such as 09:15.` | `B5` debe ser `HH:mm`, sin segundos. |
| `Config!B6 has an invalid email: ...` | Algún correo de `B6` no es válido. |
| `Duplicate employee ID: ...` | Hay un ID repetido en `Employees`. |
| `Employees!D (Country) is missing.` | Falta la columna `Country`. Ejecuta `setupAttendanceTracker`. |
| `Employees row N: Country must be MX or CO.` | País no reconocido en `Employees!D`. |
| `Vacations row N: ...` | Una fila `APPROVED` tiene datos inválidos (nombre vacío, fechas con otro formato o inicio posterior al fin). |
| `Holidays row N: date must use YYYY-MM-DD.` / `Country must be MX, CO or ALL.` | Fila inválida en `Holidays`. Corrígela. |
| `Holiday calendar for X is not available...` | La cuenta no ve el calendario de festivos del país. En Google Calendar agrégalo desde *Otros calendarios → Explorar calendarios de interés → Festivos* y repite `importHolidays`. |
| `Another attendance execution is running.` | Hay otra ejecución en curso. Espera y vuelve a intentarlo. |
| Error de la API de Chat o de autorización | Revisa que el servicio `Chat` v1 esté agregado, que aceptaste los permisos y que tu cuenta pertenece al espacio. |

Los errores del aviso de Chat están en la sección 15.

---

## 12. Limitaciones conocidas

- La ejecución diaria omite sábados y domingos. `runAttendanceForDate` los procesa, pero no escribe la pestaña del mes.
- Los festivos salen de calendarios públicos de Google. Hay que revisarlos y completarlos a mano si falta alguno.
- `Vacations`, `WhatsApp` y las pestañas mensuales relacionan al empleado por nombre. Si cambia un nombre, aparecerá una fila nueva y la anterior quedará sin actualizar.
- Los empleados inactivos no generan fila, aunque escriban su entrada.
- Quienes escriben sin estar en `Employees` aparecen en `Attendance` (en azul), pero no en las pestañas mensuales ni en el dashboard.
- Las listas desplegables de `Employees`, `Vacations`, `Holidays` y `WhatsApp` cubren hasta la fila 1000.
- `Attendance` se reescribe en cada ejecución (se borra el contenido y se vuelve a escribir). Si necesitas recuperar datos, usa *Archivo → Historial de versiones* en Google Sheets.
- El ID del Sheet y el espacio de Chat están escritos en el código. Para usarlo con otro archivo o espacio hay que editarlos.
- El corte de hora compara al segundo, sin tolerancia.
- El aviso de Chat se publica en el grupo del webhook, no en el espacio de `Config!B4`.

---

## 13. Estado de las pruebas

`v10.js` pasa la verificación de sintaxis de Node (`node --check`). El lector de fechas de la pestaña `WhatsApp` se probó con ejemplos locales. **El resto no se ha ejecutado contra Google Chat ni en Apps Script real** (aviso por webhook, festivos, dashboard). Antes de usarlo en producción, pruébalo en una copia del Google Sheet con una fecha pasada que conozcas y compara el resultado con el chat.

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
| `v10.js` | `runDailyAttendance` (hoy, sin fines de semana) y `runAttendanceForDate` (fecha de `B2`). País por empleado y pestaña `Holidays` con estado `HOLIDAY`. Dashboard. Pestaña `WhatsApp` y aviso diario al grupo de Chat por webhook (`postDailyNotice`, con su propio activador). Empleados exentos en `Config!B7`. Columna `Area` en `Employees` (lista en `Config!C`) y área en la tablas de llegadas tarde, sin registro y vacaciones del correo (con país y etiquetas de color), que ahora incluye la sección "De vacaciones". |

---

## 15. Aviso diario en Google Chat

La función `postDailyNotice` publica un mensaje en el grupo de asistencia con tres datos:

- Quién atiende los mensajes de WhatsApp hoy.
- Quién está de vacaciones hoy.
- Quién empieza vacaciones dentro de 3 días hábiles (lunes a viernes).

Ejemplo:

```
*Avisos del dia - lunes 5 de enero de 2026*

*Mensajes de WhatsApp hoy*
• González Rodríguez Joel

*De vacaciones hoy*
• Ana López (hasta el 14/01/2026)

*Proximas vacaciones (inician el jueves 8 de enero de 2026)*
• Luis Pérez (del 08/01/2026 al 12/01/2026)
```

Es texto plano con el formato básico de Chat: lo que va entre `*asteriscos*` sale en negrita.

Reglas:

- Es independiente de la asistencia: `runDailyAttendance` y `runAttendanceForDate` no publican nada. Tiene su propio activador.
- No hace nada en sábado ni domingo.
- Si no hay asignación de WhatsApp ni vacaciones, no se publica nada.
- Solo cuentan las vacaciones `APPROVED` de la pestaña `Vacations`.
- Una vacación se anuncia una vez. Si empieza en fin de semana, se anuncia como si empezara el lunes. No considera la pestaña `Holidays` al contar los 3 días.
- El mensaje no empieza con un saludo, para que `runAttendance` no lo confunda con una entrada.
- Si el envío falla, la ejecución termina con el error en el registro.
- Para cambiar la anticipación, edita `VACATION_NOTICE_DAYS_AHEAD`.

### Pestaña `WhatsApp`

La crea `setupAttendanceTracker`. Columnas: `NOMBRE` y `FECHA`.

- `NOMBRE` tiene una lista desplegable con los nombres de `Employees`. La validación solo advierte, no bloquea.
- `FECHA` tiene selector de calendario (doble clic en la celda) y se muestra como `5 de enero de 2026`. El script también lee fechas escritas como texto (`2026-01-05`, `05/01/2026`, `5 de enero de 2026`).
- Si hay varias filas con la misma fecha, salen todos los nombres.
- Una fila con fecha ilegible o sin nombre se omite y deja una advertencia en el registro.
- El nombre no se valida contra `Vacations`: si la persona asignada está de vacaciones, el aviso no lo señala.
- Si actualizas desde una versión con columna `DIA`, `setupAttendanceTracker` la elimina con su contenido.

### Configurar el webhook

Google Chat no permite publicar con la API desde una cuenta de usuario sin crear una app de Chat en Google Cloud (error `Google Chat app not found`). Por eso el aviso usa un webhook entrante.

1. En el grupo de asistencia abre *Apps e integraciones → Webhooks*, crea uno y copia la URL.
2. En Apps Script, ve a *Configuración del proyecto → Propiedades de la secuencia de comandos* y añade:
   - Propiedad: `CHAT_WEBHOOK_URL`
   - Valor: la URL completa del webhook.
3. Copia `v10.js` y `appsscript.json` al proyecto. El manifiesto debe incluir el scope `https://www.googleapis.com/auth/script.external_request`.
4. Ejecuta `postDailyNotice` a mano y acepta el permiso nuevo. Esa función publica el aviso de hoy y sirve de prueba.
5. Crea un activador diario para `postDailyNotice` (sección 10).

Notas de seguridad:

- La URL incluye una clave y un token que permiten publicar en el grupo. No la escribas en el código ni la compartas. Está solo en las propiedades del script.
- Si la URL se expone, borra el webhook en el grupo y crea uno nuevo.
- El mensaje sale con el nombre del webhook, no con el de tu cuenta.
- El webhook debe crearse en el grupo donde quieres el aviso. No usa `Config!B4`: publica donde apunte el webhook.

### Errores del aviso

| Mensaje | Causa y solución |
|---|---|
| `Missing script property CHAT_WEBHOOK_URL` | Falta la propiedad del script. Agrégala (ver arriba). |
| `Chat webhook returned HTTP 4xx` | La URL es incorrecta o el webhook fue borrado. Crea uno nuevo y actualiza la propiedad. |
| `Missing required tabs...` | Falta la pestaña `Vacations`. Ejecuta `setupAttendanceTracker`. |

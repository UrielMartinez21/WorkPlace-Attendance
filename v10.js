
/**
 * GOOGLE CHAT ATTENDANCE TRACKER
 * ========================================
 *
 * Manual attendance system using:
 * - Google Chat API
 * - Google Sheets
 *
 * FEATURES:
 * - Import employees from Google Chat
 * - Detect attendance with flexible matching (greetings,
 *   typos, emojis, extra words) plus the phrases in Config!B3
 * - Retrieve messages for a selected date
 * - Keep the earliest qualifying check-in
 * - Identify missing check-ins
 * - Preserve attendance history
 * - Avoid duplicate attendance records
 * - Apply automatic attendance colors
 * - Configure late cutoff in Google Sheets
 * - Skip check-in validation for approved vacations
 * - Filter the Attendance tab by date or employee
 * - Email report of late and missing check-ins (Config!B6)
 * - Monthly tabs (one per month) with a colored X per day:
 *   green = on time, red = late, V = vacation, F = holiday
 * - Dashboard tab: attendance and punctuality percentages, daily and
 *   monthly trends, arrival-time histogram, Mexico vs Colombia,
 *   rankings and a per-employee table, with filters by month,
 *   country and employee. Filters work live in the sheet (formulas);
 *   the data is refreshed on every attendance run.
 * - Public holidays per country (Mexico and Colombia): an employee
 *   whose country has a holiday is marked HOLIDAY, not NO CHECK-IN.
 *   Holidays are imported from Google's public holiday calendars
 *   into the Holidays tab and can be edited by hand.
 * 
 * COMMANDS IN ORDER
 * 1. testConnection            -> Verifies that the script can access your existing spreadsheet.
 * 2. setupAttendanceTracker    -> Verifies or creates the five required tabs (Config, Employees, Attendance, Vacations, Holidays) without deleting existing employee or attendance records.
 * 3. importEmployeesFromChat   -> Adds any missing members from your Google Chat attendance space. Existing employee records are preserved.
 * 4. importHolidays            -> Imports Mexico and Colombia public holidays (current and next year) into the Holidays tab. Set Employees!D (Country) by hand.
 * 5. refreshDashboard          -> Rebuilds the Dashboard data and formulas (both attendance runs already do it). Not needed to change filters: those work live from Dashboard!B3:B5.
 * 6. runAttendanceForDate      -> Manual. Processes the date typed in Config!B2 (YYYY-MM-DD, required) and updates the Attendance report. Any date, weekends included.
 * 7. runDailyAttendance        -> For the daily time trigger. Always processes today and ignores Config!B2. Does nothing on Saturdays and Sundays.
 *
 * REQUIRED SERVICE:
 * Google Chat API v1 (identifier: Chat)
 * importHolidays also uses CalendarApp and asks for permission
 * the first time.
 * The Dashboard filters need no trigger and no extra permission.
 */

// ========================================
// 1. GENERAL CONFIGURATION
// ========================================

const ATTENDANCE = {
    spreadsheetId:
        "1QEcRc1KQ-Gg1Jho7uKoX6_bBFkJw5WONzToMu3azjco",

    defaultSpace: "spaces/AAQA3dOpoiM",

    timezone: "America/Mexico_City",

    sheets: {
        config: "Config",
        employees: "Employees",
        attendance: "Attendance",
        vacations: "Vacations",
        holidays: "Holidays",
        whatsapp: "WhatsApp",
        dashboard: "Dashboard",
        dashboardData: "DashData"
    },

    // Paises soportados. Employees!D (Country) usa estas claves.
    // Todas las horas se calculan en ATTENDANCE.timezone.
    countries: {
        MX: "México",
        CO: "Colombia"
    },

    defaultCountry: "MX",

    // Calendarios publicos de Google que usa importHolidays().
    // Se prueba cada ID en orden hasta encontrar uno disponible.
    holidayCalendars: {
        MX: [
            "es.mexican#holiday@group.v.calendar.google.com",
            "en.mexican#holiday@group.v.calendar.google.com"
        ],
        CO: [
            "es.co#holiday@group.v.calendar.google.com",
            "en.co#holiday@group.v.calendar.google.com"
        ]
    },

    // Eventos de esos calendarios que no son dias no laborables
    // (fechas conmemorativas). Se compara contra el nombre y la
    // descripcion del evento, sin acentos ni mayusculas.
    holidayExcludePattern:
        "madre|padre|valent|muertos|bandera|nochebuena|" +
        "nochevieja|reyes|mujer|amor y amistad|halloween|" +
        "equinoccio|solsticio|horario|observanc|conmemor|" +
        "mother|father|valentine|christmas eve|" +
        "new year.s eve|observance",

    defaultDate: "2025-04-29",

    defaultPhrases: "presente, entrada, buenos dias",

    defaultCutoff: "09:15",

    // Frases que siempre se reconocen como asistencia,
    // ademas de las de Config!B3.
    builtInPhrases: [
        "buen dia",
        "buenos dias",
        "buenas dias",
        "buenas",
        "good morning",
        "good day",
        "presente",
        "entrada",
        "reporto entrada",
        "reporto mi entrada"
    ],

    // Palabras que pueden ir antes del saludo
    // ("Muy buenos dias", "Hola, buen dia").
    fillerWords: [
        "hola", "ola", "hey", "hello", "hi",
        "muy", "mil", "saludos"
    ],

    // Si el saludo va seguido de estas palabras no es
    // asistencia ("Buenas tardes", "Buenas noches").
    excludedNextWords: [
        "tarde", "tardes", "noche", "noches"
    ],

    monthNames: [
        "Enero", "Febrero", "Marzo", "Abril",
        "Mayo", "Junio", "Julio", "Agosto",
        "Septiembre", "Octubre", "Noviembre",
        "Diciembre"
    ],

    colors: {
        onTime: "#DCFCE7",
        onTimeText: "#166534",

        late: "#FEE2E2",
        lateText: "#991B1B",

        missing: "#FEF3C7",
        missingText: "#92400E",

        unregistered: "#DBEAFE",
        unregisteredText: "#1E40AF",

        vacation: "#E0E7FF",
        vacationText: "#3730A3",

        holiday: "#E5E7EB",
        holidayText: "#374151",

        exempt: "#CCFBF1",
        exemptText: "#115E59",

        birthday: "#FFE4E6",
        birthdayText: "#9F1239",

        // Etiquetas de color del correo. Cada area toma un color de
        // areaPalette segun su posicion en Config!C2:C50 (la 7a
        // area repite el primero). El pais usa countryColors.
        // Sin area o pais desconocido: gris.
        areaPalette: [
            { background: "#EDE9FE", text: "#5B21B6" },
            { background: "#FCE7F3", text: "#9D174D" },
            { background: "#FFEDD5", text: "#9A3412" },
            { background: "#CCFBF1", text: "#115E59" },
            { background: "#ECFCCB", text: "#3F6212" },
            { background: "#FAE8FF", text: "#86198F" }
        ],
        countryColors: {
            MX: { background: "#E0F2FE", text: "#075985" },
            CO: { background: "#FEF9C3", text: "#854D0E" }
        },
        neutralBadge: { background: "#F3F4F6", text: "#6B7280" },

        monthHeader: "#FFE599",

        header: "#17365D"
    }
};


// ========================================
// 2. CONNECT TO EXISTING SPREADSHEET
// ========================================

function getAttendanceSpreadsheet_() {
    try {
        return SpreadsheetApp.openById(
            ATTENDANCE.spreadsheetId
        );
    } catch (error) {
        throw new Error(
            "Cannot access spreadsheet: " +
            error.message
        );
    }
}


// ========================================
// 3. INITIAL SETUP
// ========================================

function setupAttendanceTracker() {
    const ss = getAttendanceSpreadsheet_();

    // CONFIG
    const config = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.config
    );

    // Forzar formato texto en cada ejecución para
    // evitar que Sheets convierta la fecha (B2) o
    // la hora (B5) a otro tipo de dato.
    config.getRange("B2:B7")
        .setNumberFormat("@");

    if (config.getLastRow() === 0) {
        config.getRange("A1:B5").setValues([
            ["Setting", "Value"],
            ["Attendance Date", ATTENDANCE.defaultDate],
            ["Valid Phrases", ATTENDANCE.defaultPhrases],
            ["Chat Space", ATTENDANCE.defaultSpace],
            ["Late After", ATTENDANCE.defaultCutoff]
        ]);

    } else {
        // Preserve existing configuration.
        // Only initialize the new cutoff field.
        if (!config.getRange("A5").getValue()) {
            config.getRange("A5")
                .setValue("Late After");
        }

        if (!config.getRange("B5").getValue()) {
            config.getRange("B5")
                .setNumberFormat("@")
                .setValue(ATTENDANCE.defaultCutoff);
        }
    }

    // Destinatarios del reporte por correo (opcional).
    // Se deja vacio: sin correos no se envia nada.
    if (!config.getRange("A6").getValue()) {
        config.getRange("A6")
            .setValue("Report Emails");
    }

    // Empleados exentos de reportar asistencia (opcional).
    // Nombres como en Employees, separados por coma o punto y
    // coma. Vacio = nadie esta exento.
    if (!config.getRange("A7").getValue()) {
        config.getRange("A7")
            .setValue("Exempt Employees");
    }

    // Lista de areas (Config!C2:C50). Alimenta la lista
    // desplegable de Employees!E. Se escriben a mano, una por
    // celda. Solo se agrega el encabezado.
    if (!config.getRange("C1").getValue()) {
        config.getRange("C1").setValue("Areas");
    }

    // EMPLOYEES
    const employees = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.employees
    );

    if (employees.getLastRow() === 0) {
        employees.getRange("A:A")
            .setNumberFormat("@");

        employees.getRange("A1:F1").setValues([[
            "Google Chat User ID",
            "Employee Name",
            "Active",
            "Country",
            "Area",
            "Birthday"
        ]]);
    }

    // Hojas creadas con versiones anteriores no tienen las
    // columnas Area y Birthday. Se agregan sin tocar los datos.
    if (employees.getMaxColumns() < 6) {
        employees.insertColumnsAfter(
            employees.getMaxColumns(),
            6 - employees.getMaxColumns()
        );
    }

    if (!employees.getRange("E1").getValue()) {
        employees.getRange("E1").setValue("Area");
    }

    if (!employees.getRange("F1").getValue()) {
        employees.getRange("F1").setValue("Birthday");
    }

    // Cumpleanos: la validacion de tipo fecha muestra el selector
    // de calendario al hacer doble clic. Solo cuentan el mes y el
    // dia; el ano puede ser cualquiera. Solo advierte, no bloquea.
    employees.getRange("F2:F1000")
        .setNumberFormat("yyyy-mm-dd")
        .setDataValidation(
            SpreadsheetApp.newDataValidation()
                .requireDate()
                .setAllowInvalid(true)
                .build()
        );

    // Area: lista tomada de Config!C2:C50. Solo se usa en el
    // correo. La validacion solo advierte, no bloquea.
    employees.getRange("E2:E1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(
                config.getRange("C2:C50"),
                true
            )
            .setAllowInvalid(true)
            .build()
    );

    // Hojas creadas con versiones anteriores no tienen la
    // columna Country. Se agrega sin tocar los datos.
    if (!employees.getRange("D1").getValue()) {
        employees.getRange("D1").setValue("Country");
    }

    // Pais: MX o CO. Define que festivos aplican al empleado.
    employees.getRange("D2:D1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInList(
                Object.keys(ATTENDANCE.countries),
                true
            )
            .setAllowInvalid(false)
            .build()
    );

    // ATTENDANCE
    const report = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.attendance
    );

    if (report.getLastRow() === 0) {
        report.getRange("A:F")
            .setNumberFormat("@");

        report.getRange("A1:F1").setValues([[
            "Date",
            "Google Chat User ID",
            "Employee",
            "Status",
            "First Check-in",
            "Message ID"
        ]]);
    }

    // VACATIONS
    const vacations = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.vacations
    );

    if (vacations.getLastRow() === 0) {
        vacations.getRange("A:D")
            .setNumberFormat("@");

        vacations.getRange("A1:D1").setValues([[
            "Employee",
            "Start Date",
            "End Date",
            "Status"
        ]]);
    }

    // Fechas de inicio y fin: la validacion de tipo fecha muestra
    // el selector de calendario al hacer doble clic. Se fuerza el
    // formato YYYY-MM-DD (el mismo que usa el script). Solo
    // advierte, para no bloquear fechas pegadas como texto.
    // Los valores que ya existen no se modifican.
    vacations.getRange("B2:C1000")
        .setNumberFormat("yyyy-mm-dd")
        .setDataValidation(
            SpreadsheetApp.newDataValidation()
                .requireDate()
                .setAllowInvalid(true)
                .build()
        );

    // Nombre del empleado: lista tomada de Employees
    // (solo advertencia, no bloquea la captura).
    vacations.getRange("A2:A1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(
                employees.getRange("B2:B1000"),
                true
            )
            .setAllowInvalid(true)
            .build()
    );

    // Estatus: lista cerrada para evitar errores de
    // captura que dejarian unas vacaciones sin aprobar.
    vacations.getRange("D2:D1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInList(
                ["APPROVED", "PENDING", "REJECTED"],
                true
            )
            .setAllowInvalid(false)
            .build()
    );

    // HOLIDAYS
    const holidays = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.holidays
    );

    if (holidays.getLastRow() === 0) {
        holidays.getRange("A:C")
            .setNumberFormat("@");

        holidays.getRange("A1:C1").setValues([[
            "Date",
            "Country",
            "Name"
        ]]);
    }

    // Pais del festivo. ALL aplica a todos los paises.
    holidays.getRange("B2:B1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInList(
                Object.keys(ATTENDANCE.countries).concat(["ALL"]),
                true
            )
            .setAllowInvalid(false)
            .build()
    );

    // WHATSAPP: quien atiende los mensajes de WhatsApp cada dia.
    // No se fuerza formato texto en la fecha: se aceptan fechas
    // reales de Sheets o texto ("2026-01-05", "5 de enero de 2026").
    const whatsapp = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.whatsapp
    );

    if (whatsapp.getLastRow() === 0) {
        whatsapp.getRange("A1:B1").setValues([[
            "NOMBRE",
            "FECHA"
        ]]);
    }

    // Versiones anteriores tenian una columna DIA (C). Ya no se
    // usa: se elimina junto con su contenido.
    if (
        whatsapp.getLastColumn() >= 3 &&
        String(whatsapp.getRange("C1").getValue()).trim() === "DIA"
    ) {
        whatsapp.deleteColumn(3);
    }

    // Fecha: la validacion de tipo fecha muestra el selector de
    // calendario al hacer doble clic en la celda. Solo advierte,
    // para no bloquear datos pegados como texto.
    whatsapp.getRange("B2:B1000")
        .setNumberFormat('d "de" mmmm "de" yyyy')
        .setDataValidation(
            SpreadsheetApp.newDataValidation()
                .requireDate()
                .setAllowInvalid(true)
                .build()
        );

    // Nombre: lista tomada de Employees (solo advertencia, no
    // bloquea la captura).
    whatsapp.getRange("A2:A1000").setDataValidation(
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(
                employees.getRange("B2:B1000"),
                true
            )
            .setAllowInvalid(true)
            .build()
    );

    // HEADER FORMATTING
    [config, employees, report, vacations, holidays, whatsapp].forEach(sheet => {
        sheet.setFrozenRows(1);

        const columns = sheet.getLastColumn();

        if (columns > 0) {
            sheet.getRange(1, 1, 1, columns)
                .setBackground(ATTENDANCE.colors.header)
                .setFontColor("#FFFFFF")
                .setFontWeight("bold");
        }
    });

    config.setColumnWidth(1, 200);
    config.setColumnWidth(2, 350);
    config.setColumnWidth(3, 200);

    employees.setColumnWidth(1, 270);
    employees.setColumnWidth(2, 300);
    employees.setColumnWidth(3, 110);
    employees.setColumnWidth(4, 110);
    employees.setColumnWidth(5, 200);
    employees.setColumnWidth(6, 130);

    report.setColumnWidth(1, 120);
    report.setColumnWidth(2, 270);
    report.setColumnWidth(3, 300);
    report.setColumnWidth(4, 220);
    report.setColumnWidth(5, 150);
    report.setColumnWidth(6, 350);

    vacations.setColumnWidth(1, 300);
    vacations.setColumnWidth(2, 130);
    vacations.setColumnWidth(3, 130);
    vacations.setColumnWidth(4, 130);

    holidays.setColumnWidth(1, 130);
    holidays.setColumnWidth(2, 110);
    holidays.setColumnWidth(3, 350);

    whatsapp.setColumnWidth(1, 330);
    whatsapp.setColumnWidth(2, 200);

    applyAttendanceFormatting();
    applyAttendanceFilter_();

    // Una pestana por mes del anio en curso.
    createMonthlySheets_(
        ss,
        Number(Utilities.formatDate(
            new Date(),
            ATTENDANCE.timezone,
            "yyyy"
        ))
    );

    refreshDashboardSafely_();

    console.log("SETUP COMPLETED");
    console.log("Existing spreadsheet connected.");
    console.log("Formatting configured.");
    console.log(ss.getUrl());
}


// ========================================
// 4. IMPORT EMPLOYEES FROM GOOGLE CHAT
// ========================================

function importEmployeesFromChat() {
    const ss = getAttendanceSpreadsheet_();

    const config = ss.getSheetByName(
        ATTENDANCE.sheets.config
    );

    const sheet = ss.getSheetByName(
        ATTENDANCE.sheets.employees
    );

    if (!config || !sheet) {
        throw new Error(
            "Run setupAttendanceTracker() first."
        );
    }

    const spaceId = String(
        config.getRange("B4").getDisplayValue()
    ).trim();

    if (!/^spaces\/[^/]+$/.test(spaceId)) {
        throw new Error("Invalid Config!B4 space ID.");
    }

    const existingIds = new Set();

    if (sheet.getLastRow() > 1) {
        const existing = sheet.getRange(
            2, 1, sheet.getLastRow() - 1, 1
        ).getDisplayValues();

        existing.forEach(row => {
            const id = String(row[0]).trim();

            if (id) existingIds.add(id);
        });
    }

    let pageToken = null;
    let totalMembers = 0;

    const newEmployees = [];

    do {
        const options = {
            pageSize: 100,
            filter: 'member.type = "HUMAN"'
        };

        if (pageToken) {
            options.pageToken = pageToken;
        }

        const response = Chat.Spaces.Members.list(
            spaceId,
            options
        );

        for (const membership of
            response.memberships || []) {

            if (
                membership.state &&
                membership.state !== "JOINED"
            ) {
                continue;
            }

            const user = membership.member || {};

            const userId = user.name;
            const userName =
                user.displayName || userId;

            if (!userId) continue;

            totalMembers++;

            if (!existingIds.has(userId)) {
                newEmployees.push([
                    userId,
                    userName,
                    true,
                    ATTENDANCE.defaultCountry
                ]);

                existingIds.add(userId);
            }
        }

        pageToken = response.nextPageToken || null;

    } while (pageToken);

    if (newEmployees.length > 0) {
        const startRow = sheet.getLastRow() + 1;

        sheet.getRange(
            startRow,
            1,
            newEmployees.length,
            4
        ).setValues(newEmployees);
    }

    console.log("EMPLOYEE IMPORT COMPLETED");
    console.log("Members found: " + totalMembers);
    console.log("New employees: " + newEmployees.length);
    console.log(
        "New employees get Country = " +
        ATTENDANCE.defaultCountry +
        ". Change Employees!D by hand where needed."
    );
    console.log("Spreadsheet: " + ss.getUrl());
}


// ========================================
// 4B. PUBLIC HOLIDAYS
// ========================================

/**
 * Importa los dias festivos de Mexico y Colombia desde los
 * calendarios publicos de Google a la pestana Holidays.
 *
 * - Sin argumento: ano en curso y el siguiente.
 * - Solo agrega fechas que faltan. No modifica ni borra filas
 *   existentes, asi que los ajustes manuales se conservan.
 *   (Una fila borrada a mano se vuelve a agregar al reimportar.)
 * - Ignora eventos que no son dias no laborables (Dia de la
 *   Madre, etc.) segun ATTENDANCE.holidayExcludePattern.
 *
 * Revisa la pestana Holidays despues de importar.
 */
function importHolidays(year) {
    const ss = getAttendanceSpreadsheet_();

    const sheet = ss.getSheetByName(
        ATTENDANCE.sheets.holidays
    );

    if (!sheet) {
        throw new Error(
            "Run setupAttendanceTracker() first."
        );
    }

    const currentYear = Number(Utilities.formatDate(
        new Date(),
        ATTENDANCE.timezone,
        "yyyy"
    ));

    const years = year
        ? [Number(year)]
        : [currentYear, currentYear + 1];

    years.forEach(y => {
        if (!Number.isInteger(y) || y < 2000 || y > 2100) {
            throw new Error("Invalid year: " + year);
        }
    });

    const exclude = new RegExp(
        ATTENDANCE.holidayExcludePattern,
        "i"
    );

    const timezone = ss.getSpreadsheetTimeZone();
    const scriptTimezone = Session.getScriptTimeZone();

    // Claves "fecha|pais" ya presentes en la hoja.
    const existing = new Set();

    if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, 2)
            .getValues()
            .forEach(row => {
                const date = toDateString_(row[0], timezone);
                const country = String(row[1]).trim().toUpperCase();

                if (date && country) {
                    existing.add(date + "|" + country);
                }
            });
    }

    const newRows = [];

    for (const country of Object.keys(ATTENDANCE.holidayCalendars)) {
        const calendar = findHolidayCalendar_(country);

        // Rango amplio; despues se filtra por ano.
        const found = new Map();

        for (const y of years) {
            const events = calendar.getEvents(
                new Date(Date.UTC(y - 1, 11, 30)),
                new Date(Date.UTC(y + 1, 0, 2))
            );

            for (const event of events) {
                if (!event.isAllDayEvent()) continue;

                const title = event.getTitle();

                const text = normalizeText_(
                    title + " " + (event.getDescription() || "")
                );

                if (exclude.test(text)) {
                    console.log(
                        "Skipped (not a holiday) | " +
                        country + " | " + title
                    );
                    continue;
                }

                // Los eventos de dia completo terminan el dia
                // siguiente (fecha final exclusiva).
                const endDate = event.getAllDayEndDate().getTime();

                for (
                    let day = event.getAllDayStartDate();
                    day.getTime() < endDate;
                    day = new Date(day.getTime() + 86400000)
                ) {
                    const date = Utilities.formatDate(
                        day,
                        scriptTimezone,
                        "yyyy-MM-dd"
                    );

                    if (!years.includes(Number(date.slice(0, 4)))) {
                        continue;
                    }

                    if (!found.has(date)) {
                        found.set(date, title);
                    }
                }
            }
        }

        for (const [date, title] of found) {
            const key = date + "|" + country;

            if (
                existing.has(key) ||
                existing.has(date + "|ALL")
            ) {
                continue;
            }

            existing.add(key);
            newRows.push([date, country, title]);
        }
    }

    if (newRows.length > 0) {
        sheet.getRange("A:C").setNumberFormat("@");

        sheet.getRange(
            sheet.getLastRow() + 1,
            1,
            newRows.length,
            3
        ).setValues(newRows);

        sheet.getRange(2, 1, sheet.getLastRow() - 1, 3)
            .sort([{ column: 1 }, { column: 2 }]);
    }

    newRows
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(row =>
            console.log(
                "Added | " + row[0] + " | " + row[1] +
                " | " + row[2]
            )
        );

    console.log("HOLIDAY IMPORT COMPLETED");
    console.log("Years: " + years.join(", "));
    console.log("Holidays added: " + newRows.length);
    console.log("Review the Holidays tab: " + ss.getUrl());
}


// Busca el primer calendario disponible de un pais.
function findHolidayCalendar_(country) {
    for (const id of ATTENDANCE.holidayCalendars[country]) {
        const calendar = CalendarApp.getCalendarById(id);

        if (calendar) {
            return calendar;
        }
    }

    throw new Error(
        "Holiday calendar for " + country + " is not available. " +
        "In Google Calendar open Other calendars > Browse calendars " +
        "of interest > Holidays, add the " +
        ATTENDANCE.countries[country] +
        " calendar and run importHolidays() again."
    );
}


// "mx", "MX", "México" -> "MX". Devuelve "" si no se reconoce.
function parseCountry_(value) {
    const key = normalizeText_(value);

    for (const code of Object.keys(ATTENDANCE.countries)) {
        if (
            key === code.toLowerCase() ||
            key === normalizeText_(ATTENDANCE.countries[code])
        ) {
            return code;
        }
    }

    return "";
}


/**
 * Lee la pestana Holidays.
 * Devuelve Map: fecha (YYYY-MM-DD) -> Map(pais -> nombre).
 * Country = ALL aplica a todos los paises.
 * Una fila con datos invalidos detiene la ejecucion.
 */
function readHolidays_(sheet, timezone) {
    const byDate = new Map();

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
        return byDate;
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();

    rows.forEach((row, index) => {
        const rowNumber = index + 2;

        const date = toDateString_(row[0], timezone);
        const countryText = String(row[1]).trim();

        if (!date && !countryText) {
            return;
        }

        if (!isValidDateString_(date)) {
            throw new Error(
                "Holidays row " + rowNumber +
                ": date must use YYYY-MM-DD."
            );
        }

        const countries =
            normalizeText_(countryText) === "all"
                ? Object.keys(ATTENDANCE.countries)
                : [parseCountry_(countryText)];

        if (!countries[0]) {
            throw new Error(
                "Holidays row " + rowNumber +
                ": Country must be " +
                Object.keys(ATTENDANCE.countries).join(", ") +
                " or ALL."
            );
        }

        const name = String(row[2]).trim() || "Festivo";

        if (!byDate.has(date)) {
            byDate.set(date, new Map());
        }

        countries.forEach(country => {
            if (!byDate.get(date).has(country)) {
                byDate.get(date).set(country, name);
            }
        });
    });

    return byDate;
}


// ========================================
// 5. MAIN ATTENDANCE FUNCTION
// ========================================

// Manual: procesa la fecha escrita en Config!B2 (YYYY-MM-DD).
// Acepta cualquier fecha, incluidos sabados y domingos.
function runAttendanceForDate() {
    runWithAttendanceLock_("date");
}


// Para el activador diario: procesa siempre el dia de hoy
// (zona ATTENDANCE.timezone), ignora Config!B2 y no hace nada
// en sabado ni domingo.
function runDailyAttendance() {
    runWithAttendanceLock_("daily");
}


function runWithAttendanceLock_(mode) {
    const lock = LockService.getScriptLock();

    if (!lock.tryLock(10000)) {
        throw new Error(
            "Another attendance execution is running."
        );
    }

    try {
        runAttendanceInternal_(mode);
    } finally {
        lock.releaseLock();
    }
}


// mode: "date" (fecha de Config!B2) o "daily" (hoy, sin fines
// de semana).
function runAttendanceInternal_(mode) {
    if (mode !== "date" && mode !== "daily") {
        throw new Error("Invalid attendance mode: " + mode);
    }

    const ss = getAttendanceSpreadsheet_();

    const config = ss.getSheetByName(
        ATTENDANCE.sheets.config
    );

    const employeeSheet = ss.getSheetByName(
        ATTENDANCE.sheets.employees
    );

    const reportSheet = ss.getSheetByName(
        ATTENDANCE.sheets.attendance
    );

    const vacationSheet = ss.getSheetByName(
        ATTENDANCE.sheets.vacations
    );

    const holidaySheet = ss.getSheetByName(
        ATTENDANCE.sheets.holidays
    );

    if (
        !config ||
        !employeeSheet ||
        !reportSheet ||
        !vacationSheet ||
        !holidaySheet
    ) {
        throw new Error(
            "Missing required tabs. " +
            "Run setupAttendanceTracker() first."
        );
    }

    // READ CONFIGURATION

    // "daily": hoy en la zona horaria configurada; Config!B2
    // se ignora. "date": la fecha que se escribio en Config!B2.
    const usingToday = mode === "daily";

    const configuredDate = String(
        config.getRange("B2").getDisplayValue()
    ).trim();

    if (!usingToday && configuredDate === "") {
        throw new Error(
            "Type the date to process in Config!B2 " +
            "(YYYY-MM-DD)."
        );
    }

    const date = usingToday
        ? Utilities.formatDate(
                new Date(),
                ATTENDANCE.timezone,
                "yyyy-MM-dd"
            )
        : configuredDate;

    const phraseText = String(
        config.getRange("B3").getDisplayValue()
    );

    const spaceId = String(
        config.getRange("B4").getDisplayValue()
    ).trim();

    // Validate date.
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        throw new Error(
            "Config!B2 must use YYYY-MM-DD."
        );
    }

    const parsedDate = new Date(
        date + "T12:00:00Z"
    );

    if (
        isNaN(parsedDate.getTime()) ||
        parsedDate.toISOString().slice(0, 10) !== date
    ) {
        throw new Error("Invalid attendance date.");
    }

    // El proceso diario no trabaja sabados ni domingos:
    // no escribe datos ni envia correo.
    if (usingToday) {
        const weekday = parsedDate.getUTCDay();

        if (weekday === 0 || weekday === 6) {
            console.log(
                "ATTENDANCE SKIPPED: " + date + " is a " +
                (weekday === 0 ? "Sunday" : "Saturday") +
                ". Nothing was processed."
            );
            return;
        }
    }

    // Validate space.
    if (!/^spaces\/[^/]+$/.test(spaceId)) {
        throw new Error("Invalid Config!B4 space ID.");
    }

    // Frases de Config!B3 mas las integradas.
    const attendancePhrases =
        buildAttendancePhrases_(phraseText);

    // Correos del reporte (Config!B6). Se validan antes
    // de procesar para fallar sin haber escrito datos.
    const reportEmails = parseReportEmails_(
        config.getRange("B6").getDisplayValue()
    );

    // Empleados exentos (Config!B7), por nombre normalizado.
    const exemptNames = parseExemptNames_(
        config.getRange("B7").getDisplayValue()
    );

    // Validate cutoff.
    const cutoff = getCutoffTime_();

    console.log("ATTENDANCE STARTED");
    console.log(
        "Date: " + date +
        (usingToday ? " (today, daily run)" : " (Config!B2)")
    );
    console.log("Space: " + spaceId);
    console.log("Late cutoff: " + cutoff);
    console.log(
        "Attendance phrases: " + attendancePhrases.length
    );

    // ========================================
    // 6. BUILD DATE RANGE
    // ========================================

    // El rango del día se calcula con la zona
    // horaria configurada (ATTENDANCE.timezone),
    // sin offset fijo, para respetar el horario
    // de verano en fechas históricas.
    const start = Utilities.parseDate(
        date + " 00:00:00",
        ATTENDANCE.timezone,
        "yyyy-MM-dd HH:mm:ss"
    );

    // Día siguiente a las 00:00 en la misma zona.
    const nextDate = Utilities.formatDate(
        new Date(parsedDate.getTime() + 86400000),
        "UTC",
        "yyyy-MM-dd"
    );

    const end = Utilities.parseDate(
        nextDate + " 00:00:00",
        ATTENDANCE.timezone,
        "yyyy-MM-dd HH:mm:ss"
    );

    const filter =
        'createTime > "' +
        new Date(start.getTime() - 1000)
            .toISOString() +
        '" AND createTime < "' +
        end.toISOString() +
        '"';

    // ========================================
    // 7. FETCH CHAT MESSAGES
    // ========================================

    let pageToken = null;
    let totalMessages = 0;

    const checkIns = new Map();

    // Mensajes que no se reconocieron como asistencia.
    const ignoredMessages = new Map();

    do {
        const options = {
            pageSize: 500,
            filter: filter
        };

        if (pageToken) {
            options.pageToken = pageToken;
        }

        const response = Chat.Spaces.Messages.list(
            spaceId,
            options
        );

        const messages = response.messages || [];

        totalMessages += messages.length;

        for (const message of messages) {
            if (
                !message.createTime ||
                !message.sender?.name
            ) {
                continue;
            }

            // Ignore bots.
            if (
                message.sender.type &&
                message.sender.type !== "HUMAN"
            ) {
                continue;
            }

            const timestamp = new Date(
                message.createTime
            );

            if (
                timestamp.getTime() < start.getTime() ||
                timestamp.getTime() >= end.getTime()
            ) {
                continue;
            }

            const messageDate = Utilities.formatDate(
                timestamp,
                ATTENDANCE.timezone,
                "yyyy-MM-dd"
            );

            if (messageDate !== date) {
                continue;
            }

            const text =
                message.text ||
                message.formattedText ||
                "";

            // Coincidencia flexible con las frases validas.
            if (!isAttendanceMessage_(text, attendancePhrases)) {
                // Guardar (maximo 3 por usuario) para el reporte
                // de mensajes ignorados.
                if (text.trim()) {
                    const list =
                        ignoredMessages.get(message.sender.name) || [];

                    if (list.length < 3) {
                        list.push(text.trim().replace(/\s+/g, " ").slice(0, 70));
                        ignoredMessages.set(message.sender.name, list);
                    }
                }

                continue;
            }

            const userId = message.sender.name;

            const existing = checkIns.get(userId);

            // Keep earliest valid check-in.
            if (
                !existing ||
                timestamp.getTime() <
                existing.timestamp.getTime()
            ) {
                checkIns.set(userId, {
                    name:
                        message.sender.displayName || userId,
                    timestamp: timestamp,
                    messageId: message.name || ""
                });
            }
        }

        pageToken = response.nextPageToken || null;

    } while (pageToken);

    console.log(
        "Messages inspected: " + totalMessages
    );

    console.log(
        "Valid unique check-ins: " + checkIns.size
    );

    // ========================================
    // 8. READ EMPLOYEE ROSTER
    // ========================================

    const roster = new Map();

    if (employeeSheet.getMaxColumns() < 4) {
        throw new Error(
            "Employees!D (Country) is missing. " +
            "Run setupAttendanceTracker() first."
        );
    }

    if (employeeSheet.getLastRow() > 1) {
        // Employees!E (Area) es opcional: si la columna aun no
        // existe, el area queda vacia ("Sin area" en el correo).
        const values = employeeSheet.getRange(
            2,
            1,
            employeeSheet.getLastRow() - 1,
            Math.min(5, employeeSheet.getMaxColumns())
        ).getDisplayValues();

        for (const [index, row] of values.entries()) {
            const userId = String(row[0]).trim();
            const name = String(row[1]).trim();

            const activeText = String(row[2])
                .trim()
                .toUpperCase();

            if (!userId) continue;

            if (roster.has(userId)) {
                throw new Error(
                    "Duplicate employee ID: " + userId
                );
            }

            const active =
                activeText === "" ||
                ["TRUE", "YES", "SI", "SÍ", "1"]
                    .includes(activeText);

            // Pais vacio = pais por defecto.
            const countryText = String(row[3]).trim();

            const country = countryText
                ? parseCountry_(countryText)
                : ATTENDANCE.defaultCountry;

            if (!country) {
                throw new Error(
                    "Employees row " + (index + 2) +
                    ": Country must be " +
                    Object.keys(ATTENDANCE.countries).join(" or ") +
                    "."
                );
            }

            roster.set(userId, {
                name: name || userId,
                active: active,
                country: country,
                area: String(row[4] || "").trim()
            });
        }
    }

    // ========================================
    // 9. GENERATE ATTENDANCE RECORDS
    // ========================================

    // Empleados con vacaciones APROBADAS que cubren
    // la fecha. Las PENDING o REJECTED se ignoran y
    // el empleado sigue el flujo normal.
    const vacationIds = getApprovedVacationIds_(
        vacationSheet,
        date,
        roster,
        ss.getSpreadsheetTimeZone()
    );

    console.log(
        "Employees on approved vacation: " +
        vacationIds.size
    );

    // Festivos de la fecha: pais -> nombre del festivo.
    const holidays = readHolidays_(
        holidaySheet,
        ss.getSpreadsheetTimeZone()
    );

    const holidaysToday = holidays.get(date) || new Map();

    if (
        !Array.from(holidays.keys()).some(
            key => key.slice(0, 4) === date.slice(0, 4)
        )
    ) {
        console.warn(
            "No holidays loaded for " + date.slice(0, 4) +
            ". Run importHolidays() or fill the Holidays tab."
        );
    }

    holidaysToday.forEach((name, country) =>
        console.log("Holiday | " + country + " | " + name)
    );

    // Resumen de festivos para el correo:
    // pais -> { country, name, count }
    const holidayInfo = new Map();

    const results = [];

    let onTime = 0;
    let late = 0;
    let noCheckIn = 0;
    let unregistered = 0;
    let vacation = 0;
    let holiday = 0;
    let exempt = 0;

    // Nombres exentos que si encontraron a un empleado activo.
    const exemptMatched = new Set();

    // Registered employees.
    for (const [userId, employee] of roster) {
        if (!employee.active) continue;

        const checkIn = checkIns.get(userId);

        let time = "";
        let status = "NO CHECK-IN";
        let messageId = "";

        const onVacation = vacationIds.has(userId);

        const holidayName = holidaysToday.get(employee.country);

        // Si reporto asistencia se conserva la hora y
        // el mensaje como referencia.
        if (checkIn) {
            time = Utilities.formatDate(
                checkIn.timestamp,
                ATTENDANCE.timezone,
                "HH:mm:ss"
            );

            messageId = checkIn.messageId;
        }

        const exemptKey = normalizeText_(employee.name);
        const isExempt = exemptNames.has(exemptKey);

        if (isExempt) {
            exemptMatched.add(exemptKey);
        }

        // Exento de reportar asistencia (Config!B7): no se evalua
        // y tiene prioridad sobre festivos y vacaciones, aunque
        // escriba su entrada (se conserva la hora de referencia).
        if (isExempt) {
            status = "EXEMPT";
            exempt++;

        // Festivo del pais del empleado: no es falta, aunque
        // escriba su entrada (se conserva la hora de referencia).
        // Tiene prioridad sobre las vacaciones.
        } else if (holidayName) {
            status = "HOLIDAY";
            holiday++;

            if (!holidayInfo.has(employee.country)) {
                holidayInfo.set(employee.country, {
                    country: employee.country,
                    name: holidayName,
                    count: 0
                });
            }

            holidayInfo.get(employee.country).count++;

        } else if (onVacation) {
            status = "ON VACATION";
            vacation++;

        } else if (checkIn) {
            status = "PRESENT";

            if (time > cutoff + ":00") {
                late++;
            } else {
                onTime++;
            }

        } else {
            noCheckIn++;

            // Ayuda a detectar frases que faltan: el empleado
            // escribio algo ese dia pero no se reconocio.
            const ignored = ignoredMessages.get(userId);

            if (ignored) {
                console.warn(
                    "NO CHECK-IN but wrote messages | " +
                    employee.name + " | " + ignored.join(" || ")
                );
            }
        }

        results.push([
            date,
            userId,
            employee.name,
            status,
            time,
            messageId
        ]);
    }

    exemptNames.forEach((original, key) => {
        if (!exemptMatched.has(key)) {
            console.warn(
                "Config!B7: no active employee named \"" +
                original + "\" in Employees."
            );
        }
    });

    // Unregistered employees.
    for (const [userId, checkIn] of checkIns) {
        if (roster.has(userId)) {
            continue;
        }

        unregistered++;

        const time = Utilities.formatDate(
            checkIn.timestamp,
            ATTENDANCE.timezone,
            "HH:mm:ss"
        );

        results.push([
            date,
            userId,
            checkIn.name,
            "PRESENT - NOT ON ROSTER",
            time,
            checkIn.messageId
        ]);
    }

    // ========================================
    // 10. PRESERVE ATTENDANCE HISTORY
    // ========================================

    const lastRow = reportSheet.getLastRow();

    let historical = [];

    if (lastRow > 1) {
        historical = reportSheet.getRange(
            2,
            1,
            lastRow - 1,
            6
        ).getDisplayValues().filter(row => {
            return String(row[0]).trim() !== date;
        });
    }

    const combined = historical.concat(results);

    combined.sort((a, b) => {
        const dateOrder = String(b[0])
            .localeCompare(String(a[0]));

        return dateOrder ||
            String(a[2]).localeCompare(String(b[2]));
    });

    // Write data only after successful API retrieval.
    if (lastRow > 1) {
        reportSheet.getRange(
            2,
            1,
            lastRow - 1,
            6
        ).clearContent();
    }

    reportSheet.getRange("A:F")
        .setNumberFormat("@");

    if (combined.length > 0) {
        reportSheet.getRange(
            2,
            1,
            combined.length,
            6
        ).setValues(combined);
    }

    // ========================================
    // 11. APPLY AUTOMATIC COLORS
    // ========================================

    applyAttendanceFormatting();
    applyAttendanceFilter_();

    // Marcar la asistencia del dia en la pestana del mes.
    updateMonthlySheet_(ss, date, results, cutoff);

    SpreadsheetApp.flush();

    // Reporte por correo: solo si hay destinatarios y
    // alguien llego tarde o no reporto.
    sendAttendanceReport_(
        reportEmails,
        date,
        results,
        cutoff,
        ss.getUrl(),
        Array.from(holidayInfo.values()),
        // ID -> { area, country } para el correo.
        new Map(
            Array.from(roster, ([userId, employee]) => [
                userId,
                { area: employee.area, country: employee.country }
            ])
        ),
        // Lista de areas (Config!C2:C50): fija el color de cada una.
        config.getRange("C2:C50").getDisplayValues()
            .map(row => String(row[0]).trim())
            .filter(Boolean),
        // Cumpleanos de hoy (y del fin de semana si es lunes).
        getBirthdayNotice_(employeeSheet, date)
    );

    // El dashboard no debe impedir que la asistencia ya
    // guardada se considere completada.
    refreshDashboardSafely_();

    // ========================================
    // 12. EXECUTION SUMMARY
    // ========================================

    console.log("ATTENDANCE COMPLETED");
    console.log("Date: " + date);
    console.log("On time: " + onTime);
    console.log("Late: " + late);
    console.log("On vacation: " + vacation);
    console.log("Holiday: " + holiday);
    console.log("Exempt: " + exempt);
    console.log("No check-in: " + noCheckIn);
    console.log("Not on roster: " + unregistered);
    console.log("Rows written: " + results.length);
    console.log("Spreadsheet: " + ss.getUrl());

    if (roster.size === 0) {
        console.warn(
            "Employee roster is empty. " +
            "Run importEmployeesFromChat()."
        );
    }
}


// ========================================
// 13. CONDITIONAL FORMATTING
// ========================================

function applyAttendanceFormatting() {
    const ss = getAttendanceSpreadsheet_();
    const sheet = ss.getSheetByName(
        ATTENDANCE.sheets.attendance
    );

    if (!sheet) {
        throw new Error("Attendance tab not found.");
    }

    // Get the deadline from Config!B5.
    const cutoff = getCutoffTime_();
    const [hour, minute] = cutoff.split(":").map(Number);

    const cutoffSeconds = hour * 3600 + minute * 60;

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
        console.log("No attendance records found.");
        return;
    }

    // Read all attendance rows.
    const range = sheet.getRange(
        2, 1, lastRow - 1, 6
    );

    const data = range.getDisplayValues();

    const backgrounds = [];
    const fontColors = [];

    const summary = {
        onTime: 0,
        late: 0,
        missing: 0,
        unregistered: 0,
        vacation: 0,
        holiday: 0
    };

    data.forEach(row => {
        const employee = row[2];
        const status = String(row[3]).trim();
        const checkIn = String(row[4]).trim();

        let background = "#FFFFFF";
        let fontColor = "#202124";

        if (status === "EXEMPT") {
            background = ATTENDANCE.colors.exempt;
            fontColor = ATTENDANCE.colors.exemptText;

        } else if (status === "HOLIDAY") {
            background = ATTENDANCE.colors.holiday;
            fontColor = ATTENDANCE.colors.holidayText;
            summary.holiday++;

        } else if (status === "ON VACATION") {
            background = ATTENDANCE.colors.vacation;
            fontColor = ATTENDANCE.colors.vacationText;
            summary.vacation++;

        } else if (status === "NO CHECK-IN") {
            background = ATTENDANCE.colors.missing;
            fontColor = ATTENDANCE.colors.missingText;
            summary.missing++;

        } else if (
            status === "PRESENT - NOT ON ROSTER"
        ) {
            background = ATTENDANCE.colors.unregistered;
            fontColor = ATTENDANCE.colors.unregisteredText;
            summary.unregistered++;

        } else if (
            status === "PRESENT" && checkIn
        ) {
            // Extract HH:mm:ss independently of
            // spreadsheet locale and cell formatting.
            const match = checkIn.match(
                /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
            );

            if (match) {
                const hours = Number(match[1]);
                const minutes = Number(match[2]);
                const seconds = Number(match[3] || 0);

                if (
                    hours <= 23 &&
                    minutes <= 59 &&
                    seconds <= 59
                ) {
                    const totalSeconds =
                        hours * 3600 +
                        minutes * 60 +
                        seconds;

                    if (totalSeconds > cutoffSeconds) {
                        // RED: Late
                        background = ATTENDANCE.colors.late;
                        fontColor = ATTENDANCE.colors.lateText;
                        summary.late++;

                        console.log(
                            "LATE: " + employee +
                            " | " + checkIn
                        );

                    } else {
                        // GREEN: On time
                        background = ATTENDANCE.colors.onTime;
                        fontColor = ATTENDANCE.colors.onTimeText;
                        summary.onTime++;
                    }
                } else {
                    console.warn(
                        "Invalid time for " + employee +
                        ": " + checkIn
                    );
                }
            } else {
                console.warn(
                    "Unrecognized time for " + employee +
                    ": " + checkIn
                );
            }
        }

        backgrounds.push(Array(6).fill(background));
        fontColors.push(Array(6).fill(fontColor));
    });

    // Remove old conditional formatting rules
    // that could override directly applied colors.
    sheet.setConditionalFormatRules([]);

    // Apply colors directly to the cells.
    range.setBackgrounds(backgrounds);
    range.setFontColors(fontColors);

    SpreadsheetApp.flush();

    console.log("ATTENDANCE COLORS APPLIED");
    console.log("Cutoff: " + cutoff);
    console.log("On time: " + summary.onTime);
    console.log("Late: " + summary.late);
    console.log("On vacation: " + summary.vacation);
    console.log("Holiday: " + summary.holiday);
    console.log("No check-in: " + summary.missing);
    console.log(
        "Not on roster: " + summary.unregistered
    );
}



// ========================================
// 14. READ AND VALIDATE CUTOFF TIME
// ========================================

function getCutoffTime_() {
    const ss = getAttendanceSpreadsheet_();

    const config = ss.getSheetByName(
        ATTENDANCE.sheets.config
    );

    if (!config) {
        throw new Error(
            "Config tab not found."
        );
    }

    const value = String(
        config.getRange("B5").getDisplayValue()
    ).trim();

    if (!/^\d{1,2}:\d{2}(:\d{2})?$/.test(value)) {
        throw new Error(
            "Config!B5 must contain a time " +
            "such as 09:15."
        );
    }

    const parts = value.split(":");

    const hour = Number(parts[0]);
    const minute = Number(parts[1]);
    const second = Number(parts[2] || 0);

    if (
        hour > 23 ||
        minute > 59 ||
        second > 59
    ) {
        throw new Error(
            "Invalid time in Config!B5."
        );
    }

    if (second !== 0) {
        throw new Error(
            "Config!B5 must use HH:mm " +
            "without seconds."
        );
    }

    return String(hour).padStart(2, "0") +
        ":" +
        String(minute).padStart(2, "0");
}


// ========================================
// 15. TEXT NORMALIZATION
// ========================================

function normalizeText_(value) {
    return String(value)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .trim()
        .replace(/^[\s.,!?¡¿]+|[\s.,!?¡¿]+$/g, "")
        .replace(/\s+/g, " ");
}


// ========================================
// 16. GET OR CREATE SHEET
// ========================================

function getOrCreateSheet_(ss, name) {
    return ss.getSheetByName(name) ||
        ss.insertSheet(name);
}


// ========================================
// 17. TEST SPREADSHEET CONNECTION
// ========================================

function testConnection() {
    const ss = getAttendanceSpreadsheet_();

    console.log("CONNECTION SUCCESSFUL");
    console.log("Spreadsheet: " + ss.getName());
    console.log("URL: " + ss.getUrl());

    ss.getSheets().forEach(sheet => {
        console.log(
            "Tab: " + sheet.getName() +
            " | Rows: " + sheet.getLastRow()
        );
    });
}


// ========================================
// 18. APPROVED VACATIONS
// ========================================

/**
 * Devuelve los IDs de empleados con vacaciones
 * APPROVED que incluyen la fecha indicada.
 *
 * Columnas de la pestana Vacations:
 * A Employee | B Start Date | C End Date | D Status
 *
 * Solo las filas APPROVED se validan y se usan.
 * El empleado se relaciona por nombre contra la
 * pestana Employees (sin acentos ni mayusculas).
 */
function getApprovedVacationIds_(
    sheet,
    date,
    roster,
    timezone
) {
    const ids = new Set();

    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
        return ids;
    }

    // Nombre normalizado -> IDs del roster.
    const idsByName = new Map();

    for (const [userId, employee] of roster) {
        const key = normalizeText_(employee.name);

        if (!idsByName.has(key)) {
            idsByName.set(key, []);
        }

        idsByName.get(key).push(userId);
    }

    const approvedValues = new Set([
        "approved",
        "aprobado",
        "aprobada"
    ]);

    const rows = sheet.getRange(
        2, 1, lastRow - 1, 4
    ).getValues();

    rows.forEach((row, index) => {
        const rowNumber = index + 2;

        if (!approvedValues.has(normalizeText_(row[3]))) {
            return;
        }

        const name = String(row[0]).trim();
        const start = toDateString_(row[1], timezone);
        const end = toDateString_(row[2], timezone);

        if (!name) {
            throw new Error(
                "Vacations row " + rowNumber +
                ": approved row without employee."
            );
        }

        if (
            !isValidDateString_(start) ||
            !isValidDateString_(end)
        ) {
            throw new Error(
                "Vacations row " + rowNumber +
                ": dates must use YYYY-MM-DD."
            );
        }

        if (start > end) {
            throw new Error(
                "Vacations row " + rowNumber +
                ": Start Date is after End Date."
            );
        }

        // Formato YYYY-MM-DD: la comparacion de texto
        // equivale a comparar fechas.
        if (date < start || date > end) {
            return;
        }

        const matches = idsByName.get(
            normalizeText_(name)
        );

        if (!matches) {
            console.warn(
                "Vacations row " + rowNumber +
                ": employee not found in Employees: " +
                name
            );
            return;
        }

        matches.forEach(userId => ids.add(userId));
    });

    return ids;
}


// ========================================
// AVISO DE VACACIONES EN GOOGLE CHAT
// ========================================

// Dias habiles (lunes a viernes) de anticipacion del aviso.
const VACATION_NOTICE_DAYS_AHEAD = 3;


// Publica el aviso de hoy (WhatsApp y vacaciones) en el grupo
// del webhook. Es independiente de runDailyAttendance: tiene su
// propio activador. No hace nada en sabado ni domingo. Tambien
// sirve para autorizar el permiso y probar el mensaje.
function postDailyNotice() {
    const ss = getAttendanceSpreadsheet_();

    const vacationSheet = ss.getSheetByName(
        ATTENDANCE.sheets.vacations
    );
    const whatsappSheet = ss.getSheetByName(
        ATTENDANCE.sheets.whatsapp
    );
    const employeeSheet = ss.getSheetByName(
        ATTENDANCE.sheets.employees
    );

    if (!vacationSheet) {
        throw new Error(
            "Missing required tabs. " +
            "Run setupAttendanceTracker() first."
        );
    }

    const today = Utilities.formatDate(
        new Date(),
        ATTENDANCE.timezone,
        "yyyy-MM-dd"
    );

    const weekday = new Date(today + "T12:00:00Z").getUTCDay();

    if (weekday === 0 || weekday === 6) {
        console.log(
            "DAILY NOTICE SKIPPED: " + today + " is a " +
            (weekday === 0 ? "Sunday" : "Saturday") + "."
        );
        return;
    }

    postDailyNotice_(
        vacationSheet,
        whatsappSheet,
        employeeSheet,
        today
    );
}


// "5 de enero de 2026", "2026-01-05", "05/01/2026" o una
// fecha real de Sheets -> YYYY-MM-DD. Devuelve "" si no se
// reconoce.
function parseAssignmentDate_(value) {
    const text = toDateString_(value, ATTENDANCE.timezone);

    if (isValidDateString_(text)) {
        return text;
    }

    const normalized = normalizeText_(text);

    let day;
    let month;
    let year;

    let match = normalized.match(
        /^(\d{1,2}) de ([a-z]+) de (\d{4})$/
    );

    if (match) {
        day = Number(match[1]);
        month = ATTENDANCE.monthNames
            .map(normalizeText_)
            .indexOf(match[2]) + 1;
        year = Number(match[3]);
    } else {
        match = normalized.match(
            /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
        );

        if (!match) {
            return "";
        }

        day = Number(match[1]);
        month = Number(match[2]);
        year = Number(match[3]);
    }

    if (month < 1) {
        return "";
    }

    const result =
        year + "-" + pad2_(month) + "-" + pad2_(day);

    return isValidDateString_(result) ? result : "";
}


// Nombres asignados a WhatsApp en "date".
// Pestana WhatsApp: A NOMBRE | B FECHA.
// Las filas con fecha ilegible se ignoran con una advertencia.
function getWhatsappAssignees_(sheet, date) {
    const names = [];

    if (!sheet || sheet.getLastRow() < 2) {
        return names;
    }

    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2)
        .getValues()
        .forEach((row, index) => {
            const name = String(row[0]).trim();

            if (!name && String(row[1]).trim() === "") {
                return;
            }

            const rowDate = parseAssignmentDate_(row[1]);

            if (!name || rowDate === "") {
                console.warn(
                    "WhatsApp row " + (index + 2) +
                    ": missing name or unreadable date."
                );
                return;
            }

            if (rowDate === date && names.indexOf(name) === -1) {
                names.push(name);
            }
        });

    return names;
}


// Publica un mensaje de texto con el webhook entrante del
// espacio. La URL contiene una clave y un token, asi que no se
// escribe en el codigo: se guarda en las propiedades del script
// (Configuracion del proyecto -> Propiedades de la secuencia de
// comandos) con el nombre CHAT_WEBHOOK_URL.
function postToChatWebhook_(text) {
    const url = PropertiesService.getScriptProperties()
        .getProperty("CHAT_WEBHOOK_URL");

    if (!url) {
        throw new Error(
            "Missing script property CHAT_WEBHOOK_URL " +
            "(Chat webhook URL)."
        );
    }

    const response = UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json; charset=UTF-8",
        payload: JSON.stringify({ text: text }),
        muteHttpExceptions: true
    });

    const code = response.getResponseCode();

    if (code < 200 || code >= 300) {
        // No se registra la URL ni el cuerpo completo.
        throw new Error(
            "Chat webhook returned HTTP " + code + ": " +
            response.getContentText().slice(0, 200)
        );
    }
}


// YYYY-MM-DD + n dias habiles (lunes a viernes).
function addBusinessDays_(date, days) {
    const current = new Date(date + "T12:00:00Z");
    let remaining = days;

    while (remaining > 0) {
        current.setUTCDate(current.getUTCDate() + 1);

        const weekday = current.getUTCDay();

        if (weekday !== 0 && weekday !== 6) {
            remaining--;
        }
    }

    return current.toISOString().slice(0, 10);
}


// Si la fecha cae en sabado o domingo, la mueve al lunes.
function nextWeekday_(date) {
    const current = new Date(date + "T12:00:00Z");

    while (
        current.getUTCDay() === 0 ||
        current.getUTCDay() === 6
    ) {
        current.setUTCDate(current.getUTCDate() + 1);
    }

    return current.toISOString().slice(0, 10);
}


// Publica quien atiende WhatsApp en "date", quien esta de
// vacaciones y quien empezara vacaciones dentro de la ventana de
// aviso: desde manana hasta VACATION_NOTICE_DAYS_AHEAD dias
// habiles. Una vacacion se avisa todos los dias de la ventana
// (no solo el ultimo), asi una aprobacion tardia tambien se
// anuncia. Las vacaciones solo cuentan en filas APPROVED. Sin
// nada que avisar, no publica.
// whatsappSheet puede ser null (pestana aun no creada).
function postDailyNotice_(
    vacationSheet,
    whatsappSheet,
    employeeSheet,
    date
) {
    const lastRow = vacationSheet.getLastRow();

    const assignees = getWhatsappAssignees_(
        whatsappSheet,
        date
    );

    const approvedValues = new Set([
        "approved",
        "aprobado",
        "aprobada"
    ]);

    // Ultimo dia de la ventana de aviso.
    const windowEnd = addBusinessDays_(
        date,
        VACATION_NOTICE_DAYS_AHEAD
    );

    const today = [];
    const upcoming = [];

    const vacationRows = lastRow < 2
        ? []
        : vacationSheet.getRange(2, 1, lastRow - 1, 4)
            .getValues();

    vacationRows
        .forEach(row => {
            if (!approvedValues.has(normalizeText_(row[3]))) {
                return;
            }

            const name = String(row[0]).trim();
            const start = toDateString_(row[1], ATTENDANCE.timezone);
            const end = toDateString_(row[2], ATTENDANCE.timezone);

            if (
                !name ||
                !isValidDateString_(start) ||
                !isValidDateString_(end) ||
                start > end
            ) {
                return;
            }

            if (start <= date && date <= end) {
                today.push({ name: name, end: end });
            } else if (
                // Si empieza en fin de semana cuenta como lunes.
                nextWeekday_(start) > date &&
                nextWeekday_(start) <= windowEnd
            ) {
                upcoming.push({
                    name: name,
                    start: start,
                    end: end
                });
            }
        });

    // Cumpleanos de hoy; los lunes, tambien los del fin de semana.
    const birthdays = getBirthdayNotice_(employeeSheet, date);

    if (
        assignees.length === 0 &&
        today.length === 0 &&
        upcoming.length === 0 &&
        birthdays.length === 0
    ) {
        console.log("Daily notice skipped: nothing to report.");
        return;
    }

    const byName = (a, b) => a.name.localeCompare(b.name);

    today.sort(byName);

    // Primero las que empiezan antes.
    upcoming.sort((a, b) =>
        a.start.localeCompare(b.start) || byName(a, b)
    );

    // Formato de Chat: *negrita*, _cursiva_ y `monoespacio`. El
    // mensaje empieza con un emoji y el titulo, nunca con un
    // saludo, para que no se confunda con un registro de entrada.
    const divider = "━━━━━━━━━━━━━━━━━━━━";

    // Quita los caracteres que Chat interpreta como formato o
    // como enlace/mencion, para que un nombre no rompa el mensaje.
    const safe = text => String(text).replace(/[*_~`<>]/g, "");

    const lines = [
        "📣 *AVISOS DEL DÍA*",
        "📅 _" + formatLongDate_(date) + "_",
        divider
    ];

    if (assignees.length > 0) {
        lines.push("");
        lines.push("💬 *WhatsApp hoy*");

        assignees.forEach(name =>
            lines.push("👤 " + safe(name))
        );
    }

    if (today.length > 0) {
        lines.push("");
        lines.push("🌴 *De vacaciones hoy*   `" + today.length + "`");

        today.forEach(item =>
            lines.push(
                "▫️ " + safe(item.name) +
                "  —  _hasta el " +
                formatMonthHeader_(item.end) + "_"
            )
        );
    }

    if (upcoming.length > 0) {
        lines.push("");
        lines.push(
            "⏳ *Próximas vacaciones*   `" + upcoming.length + "`"
        );

        // "jueves 8 de enero de 2026" -> "jueves"
        upcoming.forEach(item =>
            lines.push(
                "▫️ " + safe(item.name) +
                "  —  inicia el _" +
                formatLongDate_(item.start).split(" ")[0] + "_, " +
                "del " + formatMonthHeader_(item.start) +
                " al " + formatMonthHeader_(item.end)
            )
        );
    }

    if (birthdays.length > 0) {
        lines.push("");
        lines.push(
            "🎂 *Cumpleaños " + birthdayTitle_(birthdays) + "*   `" +
            birthdays.length + "`"
        );

        birthdays.forEach(person =>
            lines.push(
                "🎉 " + safe(person.name) +
                (person.label
                    ? "  —  _cumplió el " + person.label + "_"
                    : "")
            )
        );
    }

    lines.push("");
    lines.push(divider);
    postToChatWebhook_(lines.join("\n"));

    console.log(
        "Daily notice posted. WhatsApp: " + assignees.length +
        " | Vacation today: " + today.length +
        " | Upcoming: " + upcoming.length +
        " | Birthdays: " + birthdays.length
    );
}


// ========================================
// 19. DATE HELPERS
// ========================================

// Acepta texto YYYY-MM-DD o celdas con fecha real.
function toDateString_(value, timezone) {
    if (value instanceof Date) {
        return Utilities.formatDate(
            value,
            timezone,
            "yyyy-MM-dd"
        );
    }

    return String(value).trim();
}


function isValidDateString_(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) {
        return false;
    }

    const parsed = new Date(value + "T12:00:00Z");

    return (
        !isNaN(parsed.getTime()) &&
        parsed.toISOString().slice(0, 10) === value
    );
}


// ========================================
// 20. ATTENDANCE FILTER
// ========================================

/**
 * Aplica un filtro a toda la tabla de Attendance
 * (A:F). Desde las flechas del encabezado se puede
 * filtrar por Date (columna A), Employee (columna C)
 * o cualquier otra columna, combinando filtros.
 *
 * Cada ejecucion reescribe los datos, asi que el
 * filtro se vuelve a crear para cubrir todas las
 * filas. Los criterios que estuvieran activos se
 * conservan.
 */
function applyAttendanceFilter_() {
    const ss = getAttendanceSpreadsheet_();

    const sheet = ss.getSheetByName(
        ATTENDANCE.sheets.attendance
    );

    if (!sheet) {
        throw new Error("Attendance tab not found.");
    }

    const columns = 6;

    // Al menos una fila de datos para que el rango
    // del filtro sea valido aunque la hoja este vacia.
    const rows = Math.max(sheet.getLastRow(), 2);

    // Guardar los criterios del filtro actual.
    const savedCriteria = [];
    const currentFilter = sheet.getFilter();

    if (currentFilter) {
        for (let column = 1; column <= columns; column++) {
            const criteria =
                currentFilter.getColumnFilterCriteria(column);

            if (criteria) {
                savedCriteria.push([column, criteria]);
            }
        }

        currentFilter.remove();
    }

    const filter = sheet
        .getRange(1, 1, rows, columns)
        .createFilter();

    savedCriteria.forEach(([column, criteria]) => {
        filter.setColumnFilterCriteria(column, criteria);
    });

    console.log("ATTENDANCE FILTER APPLIED");
}


// ========================================
// 21. MONTHLY ATTENDANCE TABS
// ========================================

// Cada mes tiene su pestana (por ejemplo "Noviembre 2026"):
// una fila por empleado y una columna por dia habil
// (lunes a viernes). La celda del dia lleva:
//   X verde -> asistencia a tiempo
//   X roja  -> asistencia tarde
//   V       -> vacaciones aprobadas
//   F gris  -> festivo del pais del empleado
//   vacia   -> sin registro

function getMonthSheetName_(year, month) {
    return ATTENDANCE.monthNames[month - 1] + " " + year;
}


function pad2_(number) {
    return String(number).padStart(2, "0");
}


// Dias habiles (lunes a viernes) del mes, como YYYY-MM-DD.
function getWeekdayDates_(year, month) {
    const dates = [];

    const daysInMonth = new Date(
        Date.UTC(year, month, 0)
    ).getUTCDate();

    for (let day = 1; day <= daysInMonth; day++) {
        const weekday = new Date(
            Date.UTC(year, month - 1, day)
        ).getUTCDay();

        if (weekday === 0 || weekday === 6) {
            continue;
        }

        dates.push(
            year + "-" + pad2_(month) + "-" + pad2_(day)
        );
    }

    return dates;
}


// YYYY-MM-DD -> DD/MM/YYYY (encabezado de columna).
function formatMonthHeader_(date) {
    const parts = date.split("-");

    return parts[2] + "/" + parts[1] + "/" + parts[0];
}


// Nombres de los empleados activos de la pestana Employees.
function getActiveEmployeeNames_(ss) {
    const sheet = ss.getSheetByName(
        ATTENDANCE.sheets.employees
    );

    if (!sheet || sheet.getLastRow() < 2) {
        return [];
    }

    const values = sheet.getRange(
        2, 1, sheet.getLastRow() - 1, 3
    ).getDisplayValues();

    const names = [];

    for (const row of values) {
        const userId = String(row[0]).trim();
        const name = String(row[1]).trim();

        const activeText = String(row[2])
            .trim()
            .toUpperCase();

        const active =
            activeText === "" ||
            ["TRUE", "YES", "SI", "SÍ", "1"]
                .includes(activeText);

        if (userId && active) {
            names.push(name || userId);
        }
    }

    return names;
}


// Crea las pestanas de los 12 meses si no existen.
function createMonthlySheets_(ss, year) {
    const names = getActiveEmployeeNames_(ss);

    let created = 0;

    for (let month = 1; month <= 12; month++) {
        if (!ss.getSheetByName(
            getMonthSheetName_(year, month)
        )) {
            createMonthSheet_(ss, year, month, names);
            created++;
        }
    }

    console.log(
        "Monthly tabs created for " + year + ": " + created
    );
}


function createMonthSheet_(ss, year, month, names) {
    const sheet = ss.insertSheet(
        getMonthSheetName_(year, month)
    );

    const dates = getWeekdayDates_(year, month);
    const columns = dates.length + 1;

    // Encabezados como texto para que Sheets no
    // convierta las fechas.
    const header = sheet.getRange(1, 1, 1, columns);

    header.setNumberFormat("@");

    header.setValues([
        ["Nombre"].concat(dates.map(formatMonthHeader_))
    ]);

    header
        .setBackground(ATTENDANCE.colors.monthHeader)
        .setFontColor("#000000")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");

    sheet.getRange(1, 1)
        .setHorizontalAlignment("left");

    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(1);

    sheet.setColumnWidth(1, 330);

    for (let column = 2; column <= columns; column++) {
        sheet.setColumnWidth(column, 105);
    }

    const sorted = names.slice().sort(
        (a, b) => a.localeCompare(b, "es")
    );

    appendMonthRows_(sheet, sorted, columns);

    return sheet;
}


// Agrega filas de empleados al final y les da formato.
function appendMonthRows_(sheet, names, columns) {
    if (names.length === 0) {
        return;
    }

    const startRow = sheet.getLastRow() + 1;

    sheet.getRange(startRow, 1, names.length, 1)
        .setValues(names.map(name => [name]))
        .setBackground(ATTENDANCE.colors.monthHeader);

    if (columns > 1) {
        sheet.getRange(
            startRow, 2, names.length, columns - 1
        )
            .setHorizontalAlignment("center")
            .setFontWeight("bold");
    }

    sheet.getRange(startRow, 1, names.length, columns)
        .setBorder(
            true, true, true, true, true, true,
            "#D9D9D9",
            SpreadsheetApp.BorderStyle.SOLID
        );
}


// Marca en la pestana del mes la asistencia de "date".
// results: filas [date, userId, employee, status, time, messageId]
function updateMonthlySheet_(ss, date, results, cutoff) {
    const parts = date.split("-").map(Number);
    const year = parts[0];
    const month = parts[1];

    // Solo hay columnas para lunes a viernes.
    if (!getWeekdayDates_(year, month).includes(date)) {
        console.warn(
            "Monthly tab skipped: " + date +
            " is not a weekday."
        );
        return;
    }

    const sheetName = getMonthSheetName_(year, month);

    const sheet =
        ss.getSheetByName(sheetName) ||
        createMonthSheet_(
            ss, year, month, getActiveEmployeeNames_(ss)
        );

    const columns = sheet.getLastColumn();

    // Empleados del roster que aun no estan en la pestana.
    const lastRow = sheet.getLastRow();

    const known = new Set();

    if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, 1)
            .getDisplayValues()
            .forEach(row => known.add(normalizeText_(row[0])));
    }

    const missing = [];

    for (const row of results) {
        if (row[3] === "PRESENT - NOT ON ROSTER") continue;

        const key = normalizeText_(row[2]);

        if (!known.has(key)) {
            known.add(key);
            missing.push(String(row[2]));
        }
    }

    appendMonthRows_(sheet, missing, columns);

    // Columna del dia.
    const headers = sheet.getRange(1, 1, 1, columns)
        .getDisplayValues()[0];

    const column = headers.indexOf(
        formatMonthHeader_(date)
    ) + 1;

    if (column === 0) {
        console.warn(
            "Monthly tab: column not found for " + date
        );
        return;
    }

    // Fila de cada empleado.
    const rowCount = sheet.getLastRow() - 1;

    const rowByName = new Map();

    sheet.getRange(2, 1, rowCount, 1)
        .getDisplayValues()
        .forEach((row, index) => {
            rowByName.set(normalizeText_(row[0]), index);
        });

    const range = sheet.getRange(2, column, rowCount, 1);

    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const fontColors = range.getFontColors();

    const colors = ATTENDANCE.colors;

    for (const row of results) {
        if (row[3] === "PRESENT - NOT ON ROSTER") continue;

        const index = rowByName.get(
            normalizeText_(row[2])
        );

        if (index === undefined) continue;

        if (row[3] === "EXEMPT") {
            values[index][0] = "E";
            backgrounds[index][0] = colors.exempt;
            fontColors[index][0] = colors.exemptText;

        } else if (row[3] === "HOLIDAY") {
            values[index][0] = "F";
            backgrounds[index][0] = colors.holiday;
            fontColors[index][0] = colors.holidayText;

        } else if (row[3] === "ON VACATION") {
            values[index][0] = "V";
            backgrounds[index][0] = colors.vacation;
            fontColors[index][0] = colors.vacationText;

        } else if (row[3] === "PRESENT") {
            const late = row[4] > cutoff + ":00";

            values[index][0] = "X";
            backgrounds[index][0] =
                late ? colors.late : colors.onTime;
            fontColors[index][0] =
                late ? colors.lateText : colors.onTimeText;

        } else {
            // NO CHECK-IN: celda vacia.
            values[index][0] = "";
            backgrounds[index][0] = "#FFFFFF";
            fontColors[index][0] = "#000000";
        }
    }

    range.setValues(values);
    range.setBackgrounds(backgrounds);
    range.setFontColors(fontColors);

    console.log("MONTHLY TAB UPDATED: " + sheetName);
}


// ========================================
// 22. FLEXIBLE ATTENDANCE MESSAGE DETECTION
// ========================================

// Un mensaje cuenta como asistencia cuando EMPIEZA con una
// frase valida (Config!B3 o ATTENDANCE.builtInPhrases):
//   - Ignora mayusculas, acentos, signos y emojis.
//   - Permite palabras previas ("Muy buenos dias", "Hola, buen dia").
//   - Permite texto despues ("Buen dia, reporto mi entrada").
//   - Tolera errores pequenos de escritura ("Gooday", "buenooos").
//   - Descarta "buenas tardes" / "buenas noches".
// No cuenta si la frase aparece en medio de un texto largo
// ("Que tengan un excelente dia").

// Texto -> palabras en minuscula, sin acentos ni simbolos.
function tokenizeMessage_(value) {
    return String(value)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9\s]/g, " ")
        .replace(/(.)\1{2,}/g, "$1")
        .split(/\s+/)
        .filter(Boolean);
}


// Lista unica de frases (sin espacios) a partir de
// Config!B3 y las integradas.
function buildAttendancePhrases_(configPhrases) {
    const phrases = new Set();

    String(configPhrases)
        .split(",")
        .concat(ATTENDANCE.builtInPhrases)
        .forEach(phrase => {
            const compact = tokenizeMessage_(phrase).join("");

            if (compact) {
                phrases.add(compact);
            }
        });

    return Array.from(phrases);
}


// Errores permitidos segun la longitud de la frase.
function phraseTolerance_(phrase) {
    if (phrase.length >= 10) return 2;
    if (phrase.length >= 6) return 1;
    return 0;
}


// Distancia de edicion (Levenshtein).
function editDistance_(a, b) {
    if (a === b) return 0;

    let previous = [];

    for (let j = 0; j <= b.length; j++) {
        previous.push(j);
    }

    for (let i = 1; i <= a.length; i++) {
        const current = [i];

        for (let j = 1; j <= b.length; j++) {
            const cost = a[i - 1] === b[j - 1] ? 0 : 1;

            current.push(Math.min(
                previous[j] + 1,
                current[j - 1] + 1,
                previous[j - 1] + cost
            ));
        }

        previous = current;
    }

    return previous[b.length];
}


function isAttendanceMessage_(text, phrases) {
    const tokens = tokenizeMessage_(text);

    // Saltar palabras previas al saludo.
    let start = 0;

    while (
        start < tokens.length &&
        ATTENDANCE.fillerWords.includes(tokens[start])
    ) {
        start++;
    }

    // Probar con las primeras 1 a 4 palabras.
    for (
        let count = 1;
        count <= 4 && start + count <= tokens.length;
        count++
    ) {
        const next = tokens[start + count];

        if (
            next &&
            ATTENDANCE.excludedNextWords.includes(next)
        ) {
            continue;
        }

        const candidate = tokens
            .slice(start, start + count)
            .join("");

        for (const phrase of phrases) {
            // Descartar rapido si el largo no es compatible.
            const tolerance = phraseTolerance_(phrase);

            if (
                Math.abs(candidate.length - phrase.length) >
                tolerance
            ) {
                continue;
            }

            if (editDistance_(candidate, phrase) <= tolerance) {
                return true;
            }
        }
    }

    return false;
}


// ========================================
// 23. EMAIL REPORT
// ========================================

// Config!B6: correos separados por coma o punto y coma.
// Vacio = no se envia reporte.
function parseReportEmails_(value) {
    const emails = String(value)
        .split(/[,;\s]+/)
        .map(email => email.trim())
        .filter(Boolean);

    emails.forEach(email => {
        if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email)) {
            throw new Error(
                "Config!B6 has an invalid email: " + email
            );
        }
    });

    return Array.from(new Set(emails));
}


function escapeHtml_(value) {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}


// Config!B7: nombres separados por coma, punto y coma o salto de
// linea. Devuelve un Map nombre normalizado -> nombre escrito.
function parseExemptNames_(value) {
    const names = new Map();

    String(value)
        .split(/[,;\n]+/)
        .map(name => name.trim())
        .filter(Boolean)
        .forEach(name => names.set(normalizeText_(name), name));

    return names;
}


// YYYY-MM-DD + n dias naturales (n puede ser negativo).
function addDays_(date, days) {
    const current = new Date(date + "T12:00:00Z");

    current.setUTCDate(current.getUTCDate() + days);

    return current.toISOString().slice(0, 10);
}


// Cumpleanos que se anuncian en "date": los de ese dia y, si es
// lunes, tambien los del sabado y domingo anteriores (los avisos
// diarios no corren en fin de semana). Devuelve
// [{ userId, name, date, label }] en orden cronologico y luego por
// nombre. label es "" para los de hoy y, para los del fin de
// semana, algo como "sábado 10/01".
function getBirthdayNotice_(employeeSheet, date) {
    const dates = [date];

    if (new Date(date + "T12:00:00Z").getUTCDay() === 1) {
        dates.unshift(addDays_(date, -2), addDays_(date, -1));
    }

    const result = [];

    dates.forEach(day => {
        const label = day === date
            ? ""
            : formatLongDate_(day).split(" ")[0] + " " +
                day.slice(8) + "/" + day.slice(5, 7);

        getBirthdays_(employeeSheet, day).forEach(person =>
            result.push({
                userId: person.userId,
                name: person.name,
                date: day,
                label: label
            })
        );
    });

    return result;
}


// "de hoy", "del fin de semana" o "de hoy y del fin de semana".
function birthdayTitle_(birthdays) {
    const hasToday = birthdays.some(person => !person.label);
    const hasWeekend = birthdays.some(person => person.label);

    if (hasToday && hasWeekend) return "de hoy y del fin de semana";
    if (hasWeekend) return "del fin de semana";

    return "de hoy";
}


// Empleados activos que cumplen anos en "date" (YYYY-MM-DD).
// Lee Employees!F (Birthday); solo cuentan el mes y el dia.
// Quien nacio un 29 de febrero se felicita el 28 de febrero en
// los anos no bisiestos. Una fecha ilegible se omite con una
// advertencia. Devuelve [{ userId, name }] ordenado por nombre.
function getBirthdays_(employeeSheet, date) {
    const birthdays = [];

    if (
        !employeeSheet ||
        employeeSheet.getLastRow() < 2 ||
        employeeSheet.getMaxColumns() < 6
    ) {
        return birthdays;
    }

    const year = Number(date.slice(0, 4));

    const leapYear =
        year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);

    const monthDay = date.slice(5);

    employeeSheet
        .getRange(2, 1, employeeSheet.getLastRow() - 1, 6)
        .getDisplayValues()
        .forEach((row, index) => {
            const userId = String(row[0]).trim();
            const text = String(row[5]).trim();

            if (!userId || !text) return;

            const activeText = String(row[2]).trim().toUpperCase();

            const active =
                activeText === "" ||
                ["TRUE", "YES", "SI", "SÍ", "1"].includes(activeText);

            if (!active) return;

            const birthday = parseAssignmentDate_(text);

            if (!birthday) {
                console.warn(
                    "Employees row " + (index + 2) +
                    ": unreadable Birthday: " + text
                );
                return;
            }

            const birthdayMonthDay = birthday.slice(5);

            const matches =
                birthdayMonthDay === monthDay ||
                (
                    birthdayMonthDay === "02-29" &&
                    !leapYear &&
                    monthDay === "02-28"
                );

            if (matches) {
                birthdays.push({
                    userId: userId,
                    name: String(row[1]).trim() || userId
                });
            }
        });

    return birthdays.sort((a, b) => a.name.localeCompare(b.name));
}


// Texto de retraso: "7 min", "1 h 27 min".
function formatDelay_(time, cutoff) {
    const t = String(time).split(":").map(Number);
    const c = String(cutoff).split(":").map(Number);

    const seconds =
        (t[0] * 3600 + t[1] * 60 + (t[2] || 0)) -
        (c[0] * 3600 + c[1] * 60);

    if (seconds < 60) {
        return "menos de 1 min";
    }

    const minutes = Math.floor(seconds / 60);

    if (minutes < 60) {
        return minutes + " min";
    }

    const hours = Math.floor(minutes / 60);
    const rest = minutes % 60;

    return hours + " h" + (rest ? " " + rest + " min" : "");
}


// 2026-11-12 -> "jueves 12 de noviembre de 2026"
function formatLongDate_(date) {
    const parts = date.split("-").map(Number);

    const days = [
        "domingo", "lunes", "martes", "miércoles",
        "jueves", "viernes", "sábado"
    ];

    const weekday = new Date(
        Date.UTC(parts[0], parts[1] - 1, parts[2])
    ).getUTCDay();

    return days[weekday] + " " + parts[2] + " de " +
        ATTENDANCE.monthNames[parts[1] - 1].toLowerCase() +
        " de " + parts[0];
}


// Envia la lista de empleados tarde y sin registro.
// results: filas [date, userId, employee, status, time, messageId]
function sendAttendanceReport_(
    recipients,
    date,
    results,
    cutoff,
    spreadsheetUrl,
    holidayInfo,
    employeeInfo,
    areaList,
    birthdays
) {
    if (recipients.length === 0) {
        return;
    }

    holidayInfo = holidayInfo || [];
    areaList = areaList || [];
    birthdays = birthdays || [];

    // Area y pais de cada fila por ID de usuario (Employees!E y D).
    // Sin area asignada: "Sin área".
    const infoOf = row =>
        (employeeInfo && employeeInfo.get(row[1])) || {};

    const areaOf = row => infoOf(row).area || "Sin área";

    const countryOf = row =>
        ATTENDANCE.countries[infoOf(row).country] || "—";

    const isLate = row => row[4] > cutoff + ":00";

    const present = results.filter(row => row[3] === "PRESENT");

    const late = present
        .filter(isLate)
        .sort((a, b) => String(a[4]).localeCompare(String(b[4])));

    const onTimeCount = present.length - late.length;

    const onVacation = results
        .filter(row => row[3] === "ON VACATION")
        .sort((a, b) => String(a[2]).localeCompare(String(b[2])));

    const vacationCount = onVacation.length;

    const holidayCount = results
        .filter(row => row[3] === "HOLIDAY").length;

    // "México: Día del Trabajo (3)"
    const holidayLabels = holidayInfo.map(info =>
        ATTENDANCE.countries[info.country] + ": " +
        info.name + " (" + info.count + ")"
    );

    const missing = results
        .filter(row => row[3] === "NO CHECK-IN")
        .sort((a, b) => String(a[2]).localeCompare(String(b[2])));

    if (late.length === 0 && missing.length === 0) {
        console.log("Email report skipped: nothing to report.");
        return;
    }

    const longDate = formatLongDate_(date);

    const subject =
        "Asistencia " + date + ": " +
        late.length + " tarde, " +
        missing.length + " sin registro";

    // ---------- Version texto ----------
    const lines = [
        "Reporte de asistencia - " + longDate,
        "Hora limite de entrada: " + cutoff,
        "",
        "A tiempo: " + onTimeCount +
            " | Tarde: " + late.length +
            " | Vacaciones: " + vacationCount +
            (holidayCount > 0 ? " | Festivo: " + holidayCount : "") +
            " | Sin registro: " + missing.length,
        ...(holidayLabels.length > 0
            ? ["Festivo - " + holidayLabels.join("; ")]
            : []),
        "",
        "LLEGARON TARDE (" + late.length + ")"
    ];

    late.forEach(row =>
        lines.push(
            "- " + row[2] + " | " + areaOf(row) + " | " + countryOf(row) +
            " | " + row[4] +
            " (" + formatDelay_(row[4], cutoff) + ")"
        )
    );

    lines.push("");
    lines.push("SIN REGISTRO (" + missing.length + ")");
    missing.forEach(row =>
        lines.push(
            "- " + row[2] + " | " + areaOf(row) +
            " | " + countryOf(row)
        )
    );

    // Solo si hay personas de vacaciones.
    if (vacationCount > 0) {
        lines.push("");
        lines.push("DE VACACIONES (" + vacationCount + ")");
        onVacation.forEach(row =>
            lines.push(
                "- " + row[2] + " | " + areaOf(row) +
                " | " + countryOf(row)
            )
        );
    }

    // Solo si alguien cumple anos hoy.
    if (birthdays.length > 0) {
        lines.push("");
        lines.push("CUMPLEAÑOS (" + birthdays.length + ")");
        birthdays.forEach(person => {
            const row = [date, person.userId, person.name];

            lines.push(
                "- " + person.name + " | " + areaOf(row) +
                " | " + countryOf(row) +
                (person.label ? " | " + person.label : "")
            );
        });
    }

    lines.push("");
    lines.push(spreadsheetUrl);

    // ---------- Version HTML ----------
    // Solo estilos en linea y tablas: es lo que
    // soportan Gmail y Outlook.
    const colors = ATTENDANCE.colors;
    const font = "font-family:Arial,Helvetica,sans-serif;";

    const cardWidth = holidayCount > 0 ? "20%" : "25%";

    const statCell = (label, value, background, color) =>
        '<td align="center" width="' + cardWidth +
            '" style="padding:4px;">' +
            '<table role="presentation" width="100%" cellpadding="0" ' +
                'cellspacing="0" style="background:' + background +
                ';border-radius:8px;"><tr>' +
                '<td align="center" style="padding:14px 4px;' + font + '">' +
                    '<div style="font-size:28px;font-weight:bold;color:' +
                        color + ';line-height:1;">' + value + '</div>' +
                    '<div style="font-size:12px;color:' + color +
                        ';margin-top:6px;">' + label + '</div>' +
                '</td></tr></table>' +
        '</td>';

    const cell = (content, extra) =>
        '<td style="padding:10px 14px;border-bottom:1px solid #E5E7EB;' +
            font + 'font-size:14px;color:#111827;' + (extra || "") +
            '">' + content + '</td>';

    const headCell = (label, width) =>
        '<td align="left"' + (width ? ' width="' + width + '"' : '') +
            ' style="padding:8px 14px;' +
            font + 'font-size:11px;font-weight:bold;color:#6B7280;' +
            'text-transform:uppercase;letter-spacing:0.5px;' +
            'background:#F9FAFB;border-bottom:1px solid #E5E7EB;">' +
            label + '</td>';

    const section = (title, count, background, color, headers, bodyRows) =>
        '<tr><td style="padding:22px 28px 0 28px;">' +
            '<table role="presentation" width="100%" cellpadding="0" ' +
                'cellspacing="0"><tr>' +
                '<td style="' + font + 'font-size:16px;font-weight:bold;' +
                    'color:#111827;">' + title + '</td>' +
                '<td align="right"><span style="display:inline-block;' +
                    'padding:3px 12px;border-radius:12px;background:' +
                    background + ';color:' + color + ';' + font +
                    'font-size:13px;font-weight:bold;">' + count +
                '</span></td></tr></table>' +
            '<table role="presentation" width="100%" cellpadding="0" ' +
                'cellspacing="0" style="margin-top:10px;border:1px solid ' +
                '#E5E7EB;border-radius:8px;border-collapse:separate;' +
                'overflow:hidden;">' +
                '<tr>' + headers.join("") + '</tr>' +
                bodyRows +
            '</table>' +
        '</td></tr>';

    const emptyRow = (columns) =>
        '<tr><td colspan="' + columns + '" style="padding:14px;' + font +
            'font-size:14px;color:#6B7280;">Ninguno</td></tr>';

    // Etiqueta de color (pastilla).
    const badge = (text, style) =>
        '<span style="display:inline-block;white-space:nowrap;padding:3px 10px;' +
            'border-radius:10px;background:' + style.background +
            ';color:' + style.text + ';font-size:12px;' +
            'font-weight:bold;">' + escapeHtml_(text) + '</span>';

    // Cada area toma su color segun su posicion en Config!C.
    const areaBadge = row => {
        const area = areaOf(row);
        const index = areaList.indexOf(area);

        return badge(
            area,
            index < 0
                ? colors.neutralBadge
                : colors.areaPalette[index % colors.areaPalette.length]
        );
    };

    const countryBadge = row =>
        badge(
            countryOf(row),
            colors.countryColors[infoOf(row).country] ||
                colors.neutralBadge
        );

    const stripe = index =>
        '<tr style="background:' +
            (index % 2 ? "#F9FAFB" : "#FFFFFF") + ';">';

    const lateRows = late.length === 0
        ? emptyRow(5)
        : late.map((row, index) =>
                stripe(index) +
                    cell(escapeHtml_(row[2])) +
                    cell(areaBadge(row)) +
                    cell(countryBadge(row)) +
                    cell(escapeHtml_(row[4]), "color:" + colors.lateText + ";") +
                    cell(
                        '<span style="display:inline-block;padding:2px 10px;' +
                            'border-radius:10px;background:' + colors.late +
                            ';color:' + colors.lateText + ';font-size:12px;' +
                            'font-weight:bold;">' +
                            escapeHtml_(formatDelay_(row[4], cutoff)) + '</span>'
                    ) +
                '</tr>'
            ).join("");

    const peopleRows = list => list.map((row, index) =>
        stripe(index) +
            cell(escapeHtml_(row[2])) +
            cell(areaBadge(row)) +
            cell(countryBadge(row)) +
        '</tr>'
    ).join("");

    const missingRows = missing.length === 0
        ? emptyRow(3)
        : peopleRows(missing);

    const vacationRows = peopleRows(onVacation);

    // Mismas columnas que las otras listas: area y pais salen de
    // Employees por ID. Si hay cumpleanos del fin de semana
    // (lunes) se agrega la columna "Día".
    const birthdayHasWeekend = birthdays.some(person => person.label);

    const birthdayRows = birthdays.map((person, index) => {
        const row = [date, person.userId, person.name];

        return stripe(index) +
            cell(escapeHtml_(person.name)) +
            cell(areaBadge(row)) +
            cell(countryBadge(row)) +
            (birthdayHasWeekend
                ? cell(
                    escapeHtml_(person.label || "Hoy"),
                    "color:#4B5563;font-size:13px;white-space:nowrap;"
                )
                : "") +
        '</tr>';
    }).join("");

    const generated = Utilities.formatDate(
        new Date(),
        ATTENDANCE.timezone,
        "dd/MM/yyyy HH:mm"
    );

    const html =
        '<div style="background:#F3F4F6;padding:24px 8px;">' +
        '<table role="presentation" align="center" width="100%" ' +
            'cellpadding="0" cellspacing="0" style="max-width:820px;' +
            'margin:0 auto;background:#FFFFFF;border-radius:12px;' +
            'overflow:hidden;' + font + '">' +

            // Encabezado
            '<tr><td style="background:' + colors.header +
                ';padding:26px 28px;">' +
                '<div style="' + font + 'color:#BFD3F2;font-size:12px;' +
                    'letter-spacing:1px;text-transform:uppercase;">' +
                    '📋 Reporte de asistencia</div>' +
                '<div style="' + font + 'color:#FFFFFF;font-size:22px;' +
                    'font-weight:bold;margin-top:6px;">' +
                    escapeHtml_(longDate) + '</div>' +
                '<div style="' + font + 'color:#BFD3F2;font-size:13px;' +
                    'margin-top:6px;">Hora límite de entrada: ' +
                    escapeHtml_(cutoff) + '</div>' +
            '</td></tr>' +

            // Resumen
            '<tr><td style="padding:20px 24px 0 24px;">' +
                '<table role="presentation" width="100%" cellpadding="0" ' +
                    'cellspacing="0"><tr>' +
                    statCell("✅ A tiempo", onTimeCount,
                        colors.onTime, colors.onTimeText) +
                    statCell("⏰ Tarde", late.length,
                        colors.late, colors.lateText) +
                    statCell("🌴 Vacaciones", vacationCount,
                        colors.vacation, colors.vacationText) +
                    (holidayCount > 0
                        ? statCell("🎌 Festivo", holidayCount,
                            colors.holiday, colors.holidayText)
                        : "") +
                    statCell("⚠️ Sin registro", missing.length,
                        colors.missing, colors.missingText) +
                '</tr></table>' +
            '</td></tr>' +

            // Festivos del dia (solo si hay)
            (holidayLabels.length > 0
                ? '<tr><td style="padding:14px 28px 0 28px;' + font +
                    'font-size:13px;color:' + colors.holidayText + ';">' +
                    '<span style="display:inline-block;padding:6px 12px;' +
                    'border-radius:8px;background:' + colors.holiday + ';">' +
                    '🎌 Festivo &mdash; ' +
                    escapeHtml_(holidayLabels.join("; ")) +
                    '</span></td></tr>'
                : "") +

            // Listas
            section(
                "⏰ Llegaron tarde", late.length,
                colors.late, colors.lateText,
                [
                    headCell("Empleado", "34%"),
                    headCell("Área", "21%"),
                    headCell("País", "13%"),
                    headCell("Hora", "12%"),
                    headCell("Retraso", "20%")
                ],
                lateRows
            ) +
            section(
                "⚠️ Sin registro de entrada", missing.length,
                colors.missing, colors.missingText,
                [
                    headCell("Empleado", "54%"),
                    headCell("Área", "28%"),
                    headCell("País", "18%")
                ],
                missingRows
            ) +

            // Vacaciones (solo si hay personas)
            (vacationCount > 0
                ? section(
                    "🌴 De vacaciones", vacationCount,
                    colors.vacation, colors.vacationText,
                    [
                        headCell("Empleado", "54%"),
                        headCell("Área", "28%"),
                        headCell("País", "18%")
                    ],
                    vacationRows
                )
                : "") +

            // Cumpleanos (solo si hay personas)
            (birthdays.length > 0
                ? section(
                    "🎂 Cumpleaños " + birthdayTitle_(birthdays),
                    birthdays.length,
                    colors.birthday, colors.birthdayText,
                    birthdayHasWeekend
                        ? [
                            headCell("Empleado", "38%"),
                            headCell("Área", "24%"),
                            headCell("País", "14%"),
                            headCell("Día", "24%")
                        ]
                        : [
                            headCell("Empleado", "54%"),
                            headCell("Área", "28%"),
                            headCell("País", "18%")
                        ],
                    birthdayRows
                )
                : "") +

            // Boton
            '<tr><td align="center" style="padding:28px;">' +
                '<a href="' + escapeHtml_(spreadsheetUrl) + '" ' +
                    'style="display:inline-block;padding:12px 28px;' +
                    'background:' + colors.header + ';color:#FFFFFF;' +
                    'text-decoration:none;border-radius:8px;' + font +
                    'font-size:14px;font-weight:bold;">' +
                    '📊 Abrir hoja de asistencia</a>' +
            '</td></tr>' +

            // Pie
            '<tr><td style="padding:16px 28px;background:#F9FAFB;' +
                'border-top:1px solid #E5E7EB;' + font + 'font-size:12px;' +
                'color:#6B7280;">Generado automáticamente el ' +
                escapeHtml_(generated) + '.</td></tr>' +

        '</table></div>';

    MailApp.sendEmail({
        to: recipients.join(","),
        subject: subject,
        name: "Control de asistencia",
        body: lines.join("\n"),
        htmlBody: html
    });

    console.log(
        "Email report sent to: " + recipients.join(", ")
    );
}


// ========================================
// 24. DASHBOARD
// ========================================

// Pestana "Dashboard": filtros (mes, pais, empleado), tarjetas,
// cinco graficas, dos rankings y una tabla por empleado.
//
// Los filtros funcionan desde la propia hoja: todo lo que se ve
// son formulas que leen Dashboard!B3:B5, asi que al cambiar un
// filtro se actualiza al instante, sin ejecutar nada.
//
// El script solo escribe los datos y las formulas, y lo hace en
// cada ejecucion de asistencia (o con refreshDashboard):
//   - "DashData": una fila por empleado y dia, listas para los
//     filtros y las tablas con formulas que alimentan las graficas.
//   - "Dashboard": diseno, tarjetas, rankings y detalle.
//
// Reglas de calculo:
//   - Sabados y domingos no cuentan.
//   - Vacaciones y festivos no cuentan como falta ni bajan los
//     porcentajes: se muestran aparte.
//   - % asistencia  = (a tiempo + tarde) / (a tiempo + tarde + sin registro)
//   - % puntualidad = a tiempo / (a tiempo + tarde)
//   - "Tarde" usa la hora limite de Config!B5 vigente en la
//     ultima ejecucion de asistencia.
//   - El pais sale de Employees!D (el actual, no el de la fecha).
//   - PRESENT - NOT ON ROSTER no se incluye.

const DASHBOARD_ALL = "Todos";
const DASHBOARD_DETAIL_ROWS = 100;
const DASHBOARD_RANK_ROWS = 10;
const DASHBOARD_DAILY_ROWS = 31;
const DASHBOARD_MONTH_ROWS = 12;
const DASHBOARD_HIST_FIRST = 24; // franja de 06:00 (24 * 15 min)
const DASHBOARD_HIST_LAST = 51;  // franja de 12:45


// Reconstruye los datos y las formulas del dashboard. Las
// ejecuciones de asistencia ya lo hacen; no hace falta para
// cambiar filtros.
function refreshDashboard() {
    const lock = LockService.getScriptLock();

    if (!lock.tryLock(10000)) {
        throw new Error(
            "Another attendance execution is running."
        );
    }

    try {
        refreshDashboard_();
    } finally {
        lock.releaseLock();
    }
}


function refreshDashboardSafely_() {
    try {
        refreshDashboard_();
    } catch (error) {
        console.warn(
            "Dashboard could not be refreshed: " + error.message
        );
    }
}


// A, B, ... Z, AA, AB ...
function colLetter_(number) {
    let letters = "";

    while (number > 0) {
        const remainder = (number - 1) % 26;

        letters = String.fromCharCode(65 + remainder) + letters;
        number = Math.floor((number - 1) / 26);
    }

    return letters;
}


// Filas de Attendance listas para calcular.
// { date, month, weekday, employee, country, kind, seconds }
// kind: ontime | late | absent | vacation | holiday
function readDashboardRows_(ss, cutoff) {
    const attendance = ss.getSheetByName(
        ATTENDANCE.sheets.attendance
    );

    const employees = ss.getSheetByName(
        ATTENDANCE.sheets.employees
    );

    const rows = [];

    if (!attendance || attendance.getLastRow() < 2) {
        return rows;
    }

    const countryById = new Map();

    if (
        employees &&
        employees.getLastRow() > 1 &&
        employees.getMaxColumns() >= 4
    ) {
        employees
            .getRange(2, 1, employees.getLastRow() - 1, 4)
            .getDisplayValues()
            .forEach(row => {
                const id = String(row[0]).trim();

                if (!id) return;

                const text = String(row[3]).trim();

                countryById.set(
                    id,
                    text
                        ? parseCountry_(text)
                        : ATTENDANCE.defaultCountry
                );
            });
    }

    const cutoffParts = cutoff.split(":").map(Number);
    const cutoffSeconds =
        cutoffParts[0] * 3600 + cutoffParts[1] * 60;

    attendance
        .getRange(2, 1, attendance.getLastRow() - 1, 6)
        .getDisplayValues()
        .forEach(row => {
            const date = String(row[0]).trim();

            if (!isValidDateString_(date)) return;

            const weekday = new Date(
                date + "T12:00:00Z"
            ).getUTCDay();

            if (weekday === 0 || weekday === 6) return;

            const status = String(row[3]).trim();

            let kind = "";
            let seconds = 0;

            if (status === "NO CHECK-IN") {
                kind = "absent";

            } else if (status === "ON VACATION") {
                kind = "vacation";

            } else if (status === "HOLIDAY") {
                kind = "holiday";

            } else if (status === "PRESENT") {
                const match = String(row[4]).trim().match(
                    /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/
                );

                if (!match) return;

                seconds =
                    Number(match[1]) * 3600 +
                    Number(match[2]) * 60 +
                    Number(match[3] || 0);

                kind = seconds > cutoffSeconds ? "late" : "ontime";

            } else {
                return;
            }

            rows.push({
                date: date,
                month: date.slice(0, 7),
                weekday: weekday,
                employee: String(row[2]).trim(),
                country:
                    countryById.get(String(row[1]).trim()) || "",
                kind: kind,
                seconds: seconds,
                lateMinutes: kind === "late"
                    ? Math.round((seconds - cutoffSeconds) / 60)
                    : 0
            });
        });

    return rows;
}


function uniqueSorted_(values, descending) {
    const list = Array.from(new Set(values));

    list.sort((a, b) => String(a).localeCompare(String(b), "es"));

    return descending ? list.reverse() : list;
}


function addDashboardChart_(dash, range, row, column, title, options) {
    if (!range) return;

    const builder = dash.newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setNumHeaders(1)
        .setPosition(row, column, 0, 0)
        .setOption("title", title)
        .setOption("width", 600)
        .setOption("height", 300)
        .setOption("legend", {
            position: options.legend || "bottom"
        })
        .setOption("colors", options.colors);

    if (options.stacked) {
        builder.setOption("isStacked", true);
    }

    if (options.max) {
        builder.setOption("vAxis", {
            viewWindow: { min: 0, max: options.max }
        });
    }

    dash.insertChart(builder.build());
}


function refreshDashboard_() {
    const ss = getAttendanceSpreadsheet_();
    const cutoff = getCutoffTime_();
    const colors = ATTENDANCE.colors;

    const dash = getOrCreateSheet_(
        ss, ATTENDANCE.sheets.dashboard
    );

    const data = getOrCreateSheet_(
        ss, ATTENDANCE.sheets.dashboardData
    );

    const dashRef = "'" + ATTENDANCE.sheets.dashboard + "'!";
    const dataRef = "'" + ATTENDANCE.sheets.dashboardData + "'!";

    // Celdas de los filtros.
    const FM = dashRef + "$B$3";
    const FC = dashRef + "$B$4";
    const FE = dashRef + "$B$5";

    // ---------- Datos base ----------
    const rows = readDashboardRows_(ss, cutoff).sort((a, b) =>
        a.date.localeCompare(b.date) ||
        a.employee.localeCompare(b.employee, "es")
    );

    const employees = uniqueSorted_(rows.map(row => row.employee));
    const empIndex = new Map(
        employees.map((name, index) => [name, index + 1])
    );

    const empCountry = new Map();

    rows.forEach(row => {
        if (row.country && !empCountry.has(row.employee)) {
            empCountry.set(row.employee, row.country);
        }
    });

    const dates = uniqueSorted_(rows.map(row => row.date));
    const months = uniqueSorted_(rows.map(row => row.month));

    const dayIndex = date => Math.floor(Date.UTC(
        Number(date.slice(0, 4)),
        Number(date.slice(5, 7)) - 1,
        Number(date.slice(8, 10))
    ) / 86400000);

    const monthIndex = month =>
        Number(month.slice(0, 4)) * 12 +
        Number(month.slice(5, 7)) - 1;

    // ---------- Columnas de DashData ----------
    const C = {};
    let next = 1;

    const alloc = (name, width) => {
        C[name] = next;
        next += width || 1;
    };

    // Filas base (A:N).
    ["dayIdx", "date", "month", "weekday", "employee", "country",
        "kind", "bucket", "lateMin", "flag", "flagNoMonth",
        "flagNoCountry", "monthIdx", "empIdx"].forEach(n => alloc(n));
    next++;

    // Listas.
    ["monthOptions", "employeeOptions", "countryOptions"]
        .forEach(n => alloc(n));
    next++;
    ["dateText", "dateMonth", "dateIdx", "dateFlag", "monthLabel", "monthKey",
        "eIdx", "eName", "eCountry"].forEach(n => alloc(n));
    next++;

    // Tablas con formulas.
    ["sumLabel", "sumValue"].forEach(n => alloc(n));
    next++;
    ["dailyLabel", "dailyOn", "dailyLate", "dailyAbs"]
        .forEach(n => alloc(n));
    ["dailyHelpLabel", "dailyHelp"].forEach(n => alloc(n));
    next++;
    ["monthlyLabel", "monthlyOn", "monthlyLate", "monthlyAbs"]
        .forEach(n => alloc(n));
    ["monthlyHelpLabel", "monthlyHelp"].forEach(n => alloc(n));
    next++;
    ["histLabel", "histCount"].forEach(n => alloc(n));
    next++;
    ["ctryName", "ctryAtt", "ctryPunt", "ctryOn", "ctryLate",
        "ctryAbs"].forEach(n => alloc(n));
    next++;
    ["wdLabel", "wdLate", "wdAbs"].forEach(n => alloc(n));
    next++;
    ["eOn", "eLate", "eAbs", "eVac", "eHol", "eLateMin", "eTotal",
        "eKey", "eLateScore", "eAbsScore", "eAtt"]
        .forEach(n => alloc(n));
    next++;
    ["dKey", "dName", "dCountry", "dOn", "dLate", "dAbs", "dVac",
        "dHol", "dAtt"].forEach(n => alloc(n));
    next++;
    ["rlScore", "rlName", "rlCount", "rlAvg"]
        .forEach(n => alloc(n));
    next++;
    ["raScore", "raName", "raCount"].forEach(n => alloc(n));

    const L = name => colLetter_(C[name]);

    const ddLast = Math.max(rows.length + 1, 2);
    const dateLast = Math.max(dates.length + 1, 2);
    const monthLast = Math.max(months.length + 1, 2);
    const empLast = Math.max(employees.length + 1, 2);

    const range = (name, last) =>
        "$" + L(name) + "$2:$" + L(name) + "$" + last;
    const cellOf = (name, row) => "$" + L(name) + "$" + row;

    const dd = name => range(name, ddLast);
    const dl = name => range(name, dateLast);
    const ml = name => range(name, monthLast);
    const el = name => range(name, empLast);

    const putV = (name, row, matrix) =>
        data.getRange(row, C[name], matrix.length, matrix[0].length)
            .setValues(matrix);

    const putF = (name, row, matrix) =>
        data.getRange(row, C[name], matrix.length, matrix[0].length)
            .setFormulas(matrix);

    const count = (flagName, kindText, extra) =>
        "COUNTIFS(" + (extra || "") + dd(flagName) + ",1," +
        dd("kind") + ',"' + kindText + '")';

    // ---------- Filtros (antes de reescribir nada) ----------
    const monthOptions = [DASHBOARD_ALL].concat(months.slice().reverse());
    const employeeOptions = [DASHBOARD_ALL].concat(employees);
    const countryOptions = [DASHBOARD_ALL].concat(
        Object.keys(ATTENDANCE.countries)
    );

    const currentMonth = Utilities.formatDate(
        new Date(), ATTENDANCE.timezone, "yyyy-MM"
    );

    let month = String(dash.getRange("B3").getDisplayValue()).trim();
    let country = String(dash.getRange("B4").getDisplayValue()).trim();
    let employee = String(dash.getRange("B5").getDisplayValue()).trim();

    if (!monthOptions.includes(month)) {
        month = months.includes(currentMonth)
            ? currentMonth
            : DASHBOARD_ALL;
    }

    if (!countryOptions.includes(country)) {
        country = DASHBOARD_ALL;
    }

    if (!employeeOptions.includes(employee)) {
        employee = DASHBOARD_ALL;
    }

    // ---------- DashData: filas base ----------
    data.clear();

    const textColumns = ["date", "month", "employee", "country",
        "kind", "dateText", "dateMonth", "monthLabel", "monthOptions",
        "employeeOptions", "countryOptions", "eName", "eCountry",
        "histLabel", "ctryName", "wdLabel"];

    textColumns.forEach(name =>
        data.getRange(1, C[name], Math.max(ddLast, 600), 1).setNumberFormat("@")
    );

    const headers = [
        ["dayIdx", "DayIdx"], ["date", "Date"], ["month", "Month"],
        ["weekday", "Weekday"], ["employee", "Employee"],
        ["country", "Country"], ["kind", "Kind"],
        ["bucket", "Bucket15"], ["lateMin", "LateMin"],
        ["flag", "Flag"], ["flagNoMonth", "FlagNoMonth"],
        ["flagNoCountry", "FlagNoCountry"], ["monthIdx", "MonthIdx"],
        ["empIdx", "EmpIdx"], ["monthOptions", "Months"],
        ["employeeOptions", "Employees"],
        ["countryOptions", "Countries"], ["dateText", "Dates"],
        ["dateMonth", "DateMonth"], ["dateIdx", "DateIdx"],
        ["dateFlag", "DateInMonth"],
        ["monthLabel", "MonthAsc"], ["monthKey", "MonthKey"],
        ["eIdx", "EmpIdx"], ["eName", "EmpName"],
        ["eCountry", "EmpCountry"]
    ];

    headers.forEach(([name, label]) =>
        data.getRange(1, C[name]).setValue(label)
    );

    if (rows.length > 0) {
        data.getRange(2, C.dayIdx, rows.length, 9).setValues(
            rows.map(row => [
                dayIndex(row.date),
                row.date,
                row.month,
                row.weekday,
                row.employee,
                row.country || "--",
                row.kind,
                row.kind === "ontime" || row.kind === "late"
                    ? Math.floor(row.seconds / 900)
                    : "",
                row.lateMinutes
            ])
        );

        const flag = (useMonth, useCountry, useEmployee, rowNumber) => {
            const parts = [];

            if (useMonth) {
                parts.push(
                    'OR(' + FM + '="' + DASHBOARD_ALL + '",$' +
                    L("month") + rowNumber + "=" + FM + ")"
                );
            }

            if (useCountry) {
                parts.push(
                    'OR(' + FC + '="' + DASHBOARD_ALL + '",$' +
                    L("country") + rowNumber + "=" + FC + ")"
                );
            }

            if (useEmployee) {
                parts.push(
                    'OR(' + FE + '="' + DASHBOARD_ALL + '",$' +
                    L("employee") + rowNumber + "=" + FE + ")"
                );
            }

            return "=IF(AND(" + parts.join(",") + "),1,0)";
        };

        data.getRange(2, C.flag, rows.length, 3).setFormulas(
            rows.map((row, index) => {
                const rowNumber = index + 2;

                return [
                    flag(true, true, true, rowNumber),
                    flag(false, true, true, rowNumber),
                    flag(true, false, true, rowNumber)
                ];
            })
        );

        data.getRange(2, C.monthIdx, rows.length, 2).setValues(
            rows.map(row => [
                monthIndex(row.month),
                empIndex.get(row.employee)
            ])
        );
    }

    // ---------- DashData: listas ----------
    putV("monthOptions", 2, monthOptions.map(v => [v]));
    putV("employeeOptions", 2, employeeOptions.map(v => [v]));
    putV("countryOptions", 2, countryOptions.map(v => [v]));

    if (dates.length > 0) {
        putV("dateText", 2, dates.map(date => [
            date,
            date.slice(0, 7),
            dayIndex(date)
        ]));
    }

    // 1 si la fecha pertenece al mes elegido (para contar los dias
    // sin formulas de matriz).
    if (dates.length > 0) {
        putF("dateFlag", 2, dates.map((date, index) => [
            "=IF($" + L("dateMonth") + (index + 2) + "=" + FM +
                ",1,0)"
        ]));
    }

    if (months.length > 0) {
        putV("monthLabel", 2, months.map(m => [m, monthIndex(m)]));
    }

    if (employees.length > 0) {
        putV("eIdx", 2, employees.map((name, index) => [
            index + 1,
            name,
            empCountry.get(name) || "--"
        ]));
    }

    // ---------- DashData: resumen ----------
    putV("sumLabel", 1, [
        ["Resumen"], ["ontime"], ["late"], ["absent"],
        ["vacation"], ["holiday"], ["asistencia %"],
        ["puntualidad %"]
    ]);

    putV("sumValue", 1, [["Valor"]]);

    putF("sumValue", 2, [
        "ontime", "late", "absent", "vacation", "holiday"
    ].map(kind => ["=" + count("flag", kind)]));

    const sv = row => cellOf("sumValue", row);

    putF("sumValue", 7, [
        ['=IF(' + sv(2) + "+" + sv(3) + "+" + sv(4) + '=0,"-",' +
            "ROUND((" + sv(2) + "+" + sv(3) + ")/(" + sv(2) + "+" +
            sv(3) + "+" + sv(4) + ")*100,1))"],
        ['=IF(' + sv(2) + "+" + sv(3) + '=0,"-",' +
            "ROUND(" + sv(2) + "/(" + sv(2) + "+" + sv(3) +
            ")*100,1))"]
    ]);

    // ---------- DashData: por dia ----------
    // Mes elegido: sus dias con datos. Todos: ultimos 31 dias.
    putV("dailyLabel", 1, [["Día", "A tiempo", "Tarde", "Sin registro"]]);
    putV("dailyHelpLabel", 1, [["total"], ["primero"], ["ultimo"]]);

    const dh = row => cellOf("dailyHelp", row);

    putF("dailyHelp", 1, [
        ["=COUNT(" + dl("dateIdx") + ")"],
        ['=IF(' + FM + '="' + DASHBOARD_ALL + '",MAX(1,' + dh(1) +
            "-" + (DASHBOARD_DAILY_ROWS - 1) + "),IFERROR(MATCH(" +
            FM + "," + dl("dateMonth") + ",0),1))"],
        ['=IF(' + FM + '="' + DASHBOARD_ALL + '",' + dh(1) +
            ",IFERROR(" + dh(2) + "+SUM(" +
            dl("dateFlag") + ")-1,0))"]
    ]);

    const dailyRows = [];

    for (let k = 1; k <= DASHBOARD_DAILY_ROWS; k++) {
        const position = dh(2) + "+" + (k - 1);
        const labelCell = "$" + L("dailyLabel") + (k + 1);

        const countKind = kind =>
            "=IF(" + labelCell + '="","",COUNTIFS(' +
            dd("dayIdx") + ",INDEX(" + dl("dateIdx") + "," +
            position + ")," + dd("flag") + ",1," + dd("kind") +
            ',"' + kind + '"))';

        dailyRows.push([
            "=IF(" + position + ">" + dh(3) + ',"",RIGHT(INDEX(' +
                dl("dateText") + "," + position + '),2)&"/"&MID(INDEX(' +
                dl("dateText") + "," + position + "),6,2))",
            countKind("ontime"),
            countKind("late"),
            countKind("absent")
        ]);
    }

    putF("dailyLabel", 2, dailyRows);

    // ---------- DashData: por mes ----------
    // Ultimos 12 meses con datos (respeta pais y empleado).
    putV("monthlyLabel", 1, [["Mes", "A tiempo", "Tarde", "Sin registro"]]);
    putV("monthlyHelpLabel", 1, [["total"], ["primero"]]);

    const mh = row => cellOf("monthlyHelp", row);

    putF("monthlyHelp", 1, [
        ["=COUNT(" + ml("monthKey") + ")"],
        ["=MAX(1," + mh(1) + "-" + (DASHBOARD_MONTH_ROWS - 1) + ")"]
    ]);

    const monthlyRows = [];

    for (let k = 1; k <= DASHBOARD_MONTH_ROWS; k++) {
        const position = mh(2) + "+" + (k - 1);
        const labelCell = "$" + L("monthlyLabel") + (k + 1);

        const countKind = kind =>
            "=IF(" + labelCell + '="","",COUNTIFS(' +
            dd("monthIdx") + ",INDEX(" + ml("monthKey") + "," +
            position + ")," + dd("flagNoMonth") + ",1," +
            dd("kind") + ',"' + kind + '"))';

        monthlyRows.push([
            "=IF(" + position + ">" + mh(1) + ',"",INDEX(' +
                ml("monthLabel") + "," + position + "))",
            countKind("ontime"),
            countKind("late"),
            countKind("absent")
        ]);
    }

    putF("monthlyLabel", 2, monthlyRows);

    // ---------- DashData: hora de llegada ----------
    putV("histLabel", 1, [["Hora", "Llegadas"]]);

    const histLabels = [];
    const histFormulas = [];

    for (
        let bucket = DASHBOARD_HIST_FIRST;
        bucket <= DASHBOARD_HIST_LAST;
        bucket++
    ) {
        const minutes = bucket * 15;

        histLabels.push([
            pad2_(Math.floor(minutes / 60)) + ":" + pad2_(minutes % 60)
        ]);

        histFormulas.push([
            "=COUNTIFS(" + dd("bucket") + "," + bucket + "," +
            dd("flag") + ",1," + dd("kind") + ',"ontime")+COUNTIFS(' +
            dd("bucket") + "," + bucket + "," + dd("flag") + ",1," +
            dd("kind") + ',"late")'
        ]);
    }

    putV("histLabel", 2, histLabels);
    putF("histCount", 2, histFormulas);

    // ---------- DashData: Mexico vs Colombia ----------
    // Respeta mes y empleado, no el pais.
    putV("ctryName", 1, [[
        "País", "% asistencia", "% puntualidad",
        "ontime", "late", "absent"
    ]]);

    const countryRows = Object.keys(ATTENDANCE.countries)
        .map((code, index) => {
            const row = index + 2;

            const counted = kind =>
                "=COUNTIFS(" + dd("country") + ',"' + code + '",' +
                dd("flagNoCountry") + ",1," + dd("kind") + ',"' +
                kind + '")';

            const on = cellOf("ctryOn", row);
            const late = cellOf("ctryLate", row);
            const absent = cellOf("ctryAbs", row);

            return {
                name: ATTENDANCE.countries[code],
                formulas: [
                    "=IF(" + on + "+" + late + "+" + absent +
                        "=0,0,ROUND((" + on + "+" + late + ")/(" +
                        on + "+" + late + "+" + absent + ")*100,1))",
                    "=IF(" + on + "+" + late + "=0,0,ROUND(" + on +
                        "/(" + on + "+" + late + ")*100,1))",
                    counted("ontime"),
                    counted("late"),
                    counted("absent")
                ]
            };
        });

    putV("ctryName", 2, countryRows.map(item => [item.name]));
    putF("ctryAtt", 2, countryRows.map(item => item.formulas));

    // ---------- DashData: dia de la semana ----------
    putV("wdLabel", 1, [["Día", "Tarde", "Sin registro"]]);

    const weekdayNames = ["Lun", "Mar", "Mié", "Jue", "Vie"];

    putV("wdLabel", 2, weekdayNames.map(name => [name]));

    putF("wdLate", 2, weekdayNames.map((name, index) => [
        "=COUNTIFS(" + dd("weekday") + "," + (index + 1) + "," +
            dd("flag") + ",1," + dd("kind") + ',"late")',
        "=COUNTIFS(" + dd("weekday") + "," + (index + 1) + "," +
            dd("flag") + ",1," + dd("kind") + ',"absent")'
    ]));

    // ---------- DashData: por empleado ----------
    putV("eOn", 1, [[
        "ontime", "late", "absent", "vacation", "holiday",
        "lateMinutes", "total", "key", "lateScore", "absentScore",
        "asistencia %"
    ]]);

    const employeeRows = employees.map((name, index) => {
        const row = index + 2;
        const idx = cellOf("eIdx", row);

        const counted = kind =>
            "=COUNTIFS(" + dd("empIdx") + "," + idx + "," +
            dd("flag") + ",1," + dd("kind") + ',"' + kind + '")';

        const on = cellOf("eOn", row);
        const late = cellOf("eLate", row);
        const absent = cellOf("eAbs", row);

        return [
            counted("ontime"),
            counted("late"),
            counted("absent"),
            counted("vacation"),
            counted("holiday"),
            "=SUMIFS(" + dd("lateMin") + "," + dd("empIdx") + "," +
                idx + "," + dd("flag") + ",1," + dd("kind") +
                ',"late")',
            "=SUM(" + on + ":" + cellOf("eHol", row) + ")",
            "=IF(" + cellOf("eTotal", row) + ">0," + idx + ',"")',
            "=IF(" + late + '>0,' + late + "*100000+(99999-" + idx +
                '),"")',
            "=IF(" + absent + '>0,' + absent + "*100000+(99999-" +
                idx + '),"")',
            "=IF(" + on + "+" + late + "+" + absent + '=0,"-",' +
                "ROUND((" + on + "+" + late + ")/(" + on + "+" +
                late + "+" + absent + ")*100,1))"
        ];
    });

    if (employeeRows.length > 0) {
        putF("eOn", 2, employeeRows);
    }

    // Detalle: empleados con datos, por orden alfabetico.
    putV("dKey", 1, [[
        "key", "Empleado", "País", "A tiempo", "Tarde",
        "Sin registro", "Vacaciones", "Festivo", "% asistencia"
    ]]);

    const detailRows = [];

    for (let j = 1; j <= DASHBOARD_DETAIL_ROWS; j++) {
        const key = cellOf("dKey", j + 1);

        const pick = name =>
            "=IF(" + key + '="","",INDEX(' + el(name) + "," + key + "))";

        detailRows.push([
            "=IFERROR(SMALL(" + el("eKey") + "," + j + '),"")',
            pick("eName"),
            pick("eCountry"),
            pick("eOn"),
            pick("eLate"),
            pick("eAbs"),
            pick("eVac"),
            pick("eHol"),
            pick("eAtt")
        ]);
    }

    putF("dKey", 2, detailRows);

    // Rankings: 10 primeros (desempate alfabetico).
    putV("rlScore", 1, [[
        "score", "Empleado", "Retardos", "Retraso prom. (min)"
    ]]);
    putV("raScore", 1, [["score", "Empleado", "Faltas"]]);

    const lateRank = [];
    const absentRank = [];

    for (let j = 1; j <= DASHBOARD_RANK_ROWS; j++) {
        const lateScore = cellOf("rlScore", j + 1);
        const absentScore = cellOf("raScore", j + 1);

        const lateIdx = "99999-MOD(" + lateScore + ",100000)";
        const absentIdx = "99999-MOD(" + absentScore + ",100000)";

        lateRank.push([
            "=IFERROR(LARGE(" + el("eLateScore") + "," + j + '),"")',
            "=IF(" + lateScore + '="","",INDEX(' + el("eName") + "," +
                lateIdx + "))",
            "=IF(" + lateScore + '="","",INT(' + lateScore +
                "/100000))",
            "=IF(" + lateScore + '="","",ROUND(INDEX(' +
                el("eLateMin") + "," + lateIdx + ")/INT(" +
                lateScore + "/100000),0))"
        ]);

        absentRank.push([
            "=IFERROR(LARGE(" + el("eAbsScore") + "," + j + '),"")',
            "=IF(" + absentScore + '="","",INDEX(' + el("eName") +
                "," + absentIdx + "))",
            "=IF(" + absentScore + '="","",INT(' + absentScore +
                "/100000))"
        ]);
    }

    putF("rlScore", 2, lateRank);
    putF("raScore", 2, absentRank);

    // ---------- Dashboard: filtros ----------
    const listRule = (name, size) =>
        SpreadsheetApp.newDataValidation()
            .requireValueInRange(
                data.getRange(2, C[name], size, 1),
                true
            )
            .setAllowInvalid(false)
            .build();

    dash.getRange("B3:B5").setNumberFormat("@");

    dash.getRange("A1").setValue("Dashboard de asistencia")
        .setFontSize(18).setFontWeight("bold");

    dash.getRange("A3:A5").setValues([
        ["Mes"], ["País"], ["Empleado"]
    ]).setFontWeight("bold");

    dash.getRange("B3").setValue(month)
        .setDataValidation(listRule("monthOptions", monthOptions.length));
    dash.getRange("B4").setValue(country)
        .setDataValidation(listRule("countryOptions", countryOptions.length));
    dash.getRange("B5").setValue(employee)
        .setDataValidation(
            listRule("employeeOptions", employeeOptions.length)
        );

    dash.getRange("B3:B5").setBackground("#FFF2CC");

    dash.getRange("A6").setValue(
        "Los filtros se aplican al instante. Hora límite " + cutoff +
        " (hora de México). Sábados y domingos no cuentan; " +
        "vacaciones y festivos no cuentan como falta. " +
        "Los datos se actualizan en cada ejecución de asistencia."
    ).setFontColor("#6B7280").setFontSize(9);

    // Limpiar el resto (se vuelve a escribir completo).
    dash.getRange(
        7, 1, Math.max(dash.getMaxRows() - 6, 1), 12
    ).clear();

    // Anchos: A y G llevan nombres.
    [200, 90, 90, 90, 90, 90, 170, 90, 90, 90, 90, 90]
        .forEach((width, index) =>
            dash.setColumnWidth(index + 1, width)
        );

    // ---------- Dashboard: tarjetas ----------
    const ref = (name, row) => dataRef + "$" + L(name) + "$" + row;

    const percentCard = row =>
        "=IF(ISNUMBER(" + dataRef + sv(row) + ")," + dataRef +
        sv(row) + '&"%","-")';

    const cards = [
        ["Asistencia", percentCard(7), colors.onTime, colors.onTimeText],
        ["Puntualidad", percentCard(8), colors.onTime, colors.onTimeText],
        ["Retardos", "=" + dataRef + sv(3), colors.late, colors.lateText],
        ["Faltas", "=" + dataRef + sv(4), colors.missing, colors.missingText],
        ["Vacaciones", "=" + dataRef + sv(5), colors.vacation,
            colors.vacationText],
        ["Festivos", "=" + dataRef + sv(6), colors.holiday,
            colors.holidayText]
    ];

    cards.forEach((card, index) => {
        dash.getRange(8, index + 1, 2, 1)
            .setBackground(card[2])
            .setFontColor(card[3])
            .setHorizontalAlignment("center");

        dash.getRange(8, index + 1).setValue(card[0])
            .setFontSize(10);

        dash.getRange(9, index + 1).setFormula(card[1])
            .setFontSize(20).setFontWeight("bold");
    });

    dash.setRowHeight(9, 40);

    // ---------- Dashboard: graficas ----------
    dash.getCharts().forEach(chart => dash.removeChart(chart));

    const triple = ["#16A34A", "#DC2626", "#F59E0B"];

    const tableRange = (name, rowsCount, columns) =>
        data.getRange(1, C[name], rowsCount + 1, columns);

    addDashboardChart_(dash,
        tableRange("dailyLabel", DASHBOARD_DAILY_ROWS, 4), 11, 1,
        "Asistencia por día (mes elegido o últimos 31 días)",
        { stacked: true, colors: triple });

    addDashboardChart_(dash,
        tableRange("monthlyLabel", DASHBOARD_MONTH_ROWS, 4), 11, 7,
        "Asistencia por mes (últimos 12)",
        { stacked: true, colors: triple });

    addDashboardChart_(dash,
        tableRange("histLabel", histLabels.length, 2), 27, 1,
        "Hora de llegada, 06:00 a 12:45 (franjas de 15 min)",
        { colors: ["#2563EB"], legend: "none" });

    addDashboardChart_(dash,
        tableRange("ctryName", countryRows.length, 3), 27, 7,
        "México vs Colombia (%)",
        { colors: ["#16A34A", "#2563EB"], max: 100 });

    addDashboardChart_(dash,
        tableRange("wdLabel", weekdayNames.length, 3), 43, 1,
        "Retardos y faltas por día de la semana",
        { colors: ["#DC2626", "#F59E0B"] });

    // ---------- Dashboard: rankings y detalle ----------
    const header = (row, column, labels) =>
        dash.getRange(row, column, 1, labels.length)
            .setValues([labels])
            .setBackground(colors.header)
            .setFontColor("#FFFFFF")
            .setFontWeight("bold");

    const title = (row, column, text) =>
        dash.getRange(row, column).setValue(text)
            .setFontSize(13).setFontWeight("bold");

    const link = (names, firstRow, count) => {
        const matrix = [];

        for (let j = 0; j < count; j++) {
            matrix.push(names.map(name =>
                "=" + ref(name, firstRow + j)
            ));
        }

        return matrix;
    };

    title(59, 1, "Más retardos");
    header(60, 1, ["Empleado", "Retardos", "Retraso prom. (min)"]);
    dash.getRange(61, 1, DASHBOARD_RANK_ROWS, 3).setFormulas(
        link(["rlName", "rlCount", "rlAvg"], 2, DASHBOARD_RANK_ROWS)
    );

    title(59, 7, "Más faltas");
    header(60, 7, ["Empleado", "Faltas"]);
    dash.getRange(61, 7, DASHBOARD_RANK_ROWS, 2).setFormulas(
        link(["raName", "raCount"], 2, DASHBOARD_RANK_ROWS)
    );

    title(72, 1, "Detalle por empleado");
    header(73, 1, [
        "Empleado", "País", "A tiempo", "Tarde",
        "Sin registro", "Vacaciones", "Festivo", "% asistencia"
    ]);

    dash.getRange(74, 1, DASHBOARD_DETAIL_ROWS, 8).setFormulas(
        link(
            ["dName", "dCountry", "dOn", "dLate", "dAbs", "dVac",
                "dHol", "dAtt"],
            2,
            DASHBOARD_DETAIL_ROWS
        )
    );

    dash.getRange(74, 2, DASHBOARD_DETAIL_ROWS, 7)
        .setHorizontalAlignment("center");

    const attendanceRange = dash.getRange(
        74, 8, DASHBOARD_DETAIL_ROWS, 1
    );

    attendanceRange.setNumberFormat('0.0"%"');

    // Semaforo del % de asistencia.
    dash.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThanOrEqualTo(95)
            .setBackground(colors.onTime)
            .setFontColor(colors.onTimeText)
            .setRanges([attendanceRange])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(85, 94.99)
            .setBackground(colors.missing)
            .setFontColor(colors.missingText)
            .setRanges([attendanceRange])
            .build(),
        SpreadsheetApp.newConditionalFormatRule()
            .whenNumberLessThan(85)
            .setBackground(colors.late)
            .setFontColor(colors.lateText)
            .setRanges([attendanceRange])
            .build()
    ]);

    if (employees.length > DASHBOARD_DETAIL_ROWS) {
        console.warn(
            "Dashboard detail shows the first " +
            DASHBOARD_DETAIL_ROWS + " of " + employees.length +
            " employees."
        );
    }

    dash.setFrozenRows(0);

    console.log("DASHBOARD UPDATED");
    console.log("Rows written: " + rows.length);
    console.log(
        "Filters kept: " + month + " | " + country + " | " + employee
    );
}


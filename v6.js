
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
 * - Detect attendance using configured phrases
 * - Retrieve messages for a selected date
 * - Keep the earliest qualifying check-in
 * - Identify missing check-ins
 * - Preserve attendance history
 * - Avoid duplicate attendance records
 * - Apply automatic attendance colors
 * - Configure late cutoff in Google Sheets
 * - Skip check-in validation for approved vacations
 * - Filter the Attendance tab by date or employee
 * - Monthly tabs (one per month) with a colored X per day:
 *   green = on time, red = late, V = vacation
 * 
 * COMMANDS IN ORDER
 * 1. testConnection            -> Verifies that the script can access your existing spreadsheet.
 * 2. setupAttendanceTracker    -> Verifies or creates the four required tabs (Config, Employees, Attendance, Vacations) without deleting existing employee or attendance records.
 * 3. importEmployeesFromChat   -> Adds any missing members from your Google Chat attendance space. Existing employee records are preserved.
 * 4. runAttendance             -> Processes the date in Config!B2 (or today if B2 is empty) and updates the Attendance report.
 *
 * REQUIRED SERVICE:
 * Google Chat API v1 (identifier: Chat)
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
        vacations: "Vacations"
    },

    defaultDate: "2025-04-29",

    defaultPhrases: "presente, entrada, buenos dias",

    defaultCutoff: "09:15",

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
    config.getRange("B2:B5")
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

    // EMPLOYEES
    const employees = getOrCreateSheet_(
        ss,
        ATTENDANCE.sheets.employees
    );

    if (employees.getLastRow() === 0) {
        employees.getRange("A:A")
            .setNumberFormat("@");

        employees.getRange("A1:C1").setValues([[
            "Google Chat User ID",
            "Employee Name",
            "Active"
        ]]);
    }

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

    // HEADER FORMATTING
    [config, employees, report, vacations].forEach(sheet => {
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

    employees.setColumnWidth(1, 270);
    employees.setColumnWidth(2, 300);
    employees.setColumnWidth(3, 110);

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
                    true
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
            3
        ).setValues(newEmployees);
    }

    console.log("EMPLOYEE IMPORT COMPLETED");
    console.log("Members found: " + totalMembers);
    console.log("New employees: " + newEmployees.length);
    console.log("Spreadsheet: " + ss.getUrl());
}


// ========================================
// 5. MAIN ATTENDANCE FUNCTION
// ========================================

function runAttendance() {
    const lock = LockService.getScriptLock();

    if (!lock.tryLock(10000)) {
        throw new Error(
            "Another attendance execution is running."
        );
    }

    try {
        runAttendanceInternal_();
    } finally {
        lock.releaseLock();
    }
}


function runAttendanceInternal_() {
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

    if (
        !config ||
        !employeeSheet ||
        !reportSheet ||
        !vacationSheet
    ) {
        throw new Error(
            "Missing required tabs. " +
            "Run setupAttendanceTracker() first."
        );
    }

    // READ CONFIGURATION

    // Si Config!B2 esta vacia se usa el dia de hoy
    // (en la zona horaria configurada).
    const configuredDate = String(
        config.getRange("B2").getDisplayValue()
    ).trim();

    const usingToday = configuredDate === "";

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
            "Config!B2 must use YYYY-MM-DD or be empty."
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

    // Validate space.
    if (!/^spaces\/[^/]+$/.test(spaceId)) {
        throw new Error("Invalid Config!B4 space ID.");
    }

    // Valid phrases come directly from Google Sheets.
    const phrases = new Set(
        phraseText
            .split(",")
            .map(normalizeText_)
            .filter(Boolean)
    );

    if (phrases.size === 0) {
        throw new Error(
            "No attendance phrases in Config!B3."
        );
    }

    // Validate cutoff.
    const cutoff = getCutoffTime_();

    console.log("ATTENDANCE STARTED");
    console.log(
        "Date: " + date +
        (usingToday ? " (today, Config!B2 is empty)" : "")
    );
    console.log("Space: " + spaceId);
    console.log("Late cutoff: " + cutoff);

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

            const text = normalizeText_(
                message.text || ""
            );

            // EXACT MATCH against Config!B3.
            if (!phrases.has(text)) {
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

    if (employeeSheet.getLastRow() > 1) {
        const values = employeeSheet.getRange(
            2,
            1,
            employeeSheet.getLastRow() - 1,
            3
        ).getDisplayValues();

        for (const row of values) {
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

            roster.set(userId, {
                name: name || userId,
                active: active
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

    const results = [];

    let onTime = 0;
    let late = 0;
    let noCheckIn = 0;
    let unregistered = 0;
    let vacation = 0;

    // Registered employees.
    for (const [userId, employee] of roster) {
        if (!employee.active) continue;

        const checkIn = checkIns.get(userId);

        let time = "";
        let status = "NO CHECK-IN";
        let messageId = "";

        const onVacation = vacationIds.has(userId);

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

        if (onVacation) {
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

    // ========================================
    // 12. EXECUTION SUMMARY
    // ========================================

    console.log("ATTENDANCE COMPLETED");
    console.log("Date: " + date);
    console.log("On time: " + onTime);
    console.log("Late: " + late);
    console.log("On vacation: " + vacation);
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
        vacation: 0
    };

    data.forEach(row => {
        const employee = row[2];
        const status = String(row[3]).trim();
        const checkIn = String(row[4]).trim();

        let background = "#FFFFFF";
        let fontColor = "#202124";

        if (status === "ON VACATION") {
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

        if (row[3] === "ON VACATION") {
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

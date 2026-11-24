
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
 * 
 * COMMANDS IN ORDER
 * 1. testConnection            -> Verifies that the script can access your existing spreadsheet.
 * 2. setupAttendanceTracker    -> Verifies or creates the three required tabs without deleting existing employee or attendance records.
 * 3. importEmployeesFromChat   -> Adds any missing members from your Google Chat attendance space. Existing employee records are preserved.
 * 4. runAttendance             -> Processes the date configured in Config!B2 and updates the Attendance report.
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
        attendance: "Attendance"
    },

    defaultDate: "2025-04-29",

    defaultPhrases: "presente, entrada, buenos dias",

    defaultCutoff: "09:15",

    colors: {
        onTime: "#DCFCE7",
        onTimeText: "#166534",

        late: "#FEE2E2",
        lateText: "#991B1B",

        missing: "#FEF3C7",
        missingText: "#92400E",

        unregistered: "#DBEAFE",
        unregisteredText: "#1E40AF",

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

    if (config.getLastRow() === 0) {
        config.getRange("B2:B5")
            .setNumberFormat("@");

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

    // HEADER FORMATTING
    [config, employees, report].forEach(sheet => {
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

    applyAttendanceFormatting();

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

    if (!config || !employeeSheet || !reportSheet) {
        throw new Error(
            "Missing required tabs. " +
            "Run setupAttendanceTracker() first."
        );
    }

    // READ CONFIGURATION

    const date = String(
        config.getRange("B2").getDisplayValue()
    ).trim();

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
    console.log("Date: " + date);
    console.log("Space: " + spaceId);
    console.log("Late cutoff: " + cutoff);

    // ========================================
    // 6. BUILD DATE RANGE
    // ========================================

    // Current Mexico City timezone: UTC-06:00.
    // This is appropriate for the configured
    // modern attendance dates.
    const start = new Date(
        date + "T00:00:00-06:00"
    );

    const end = new Date(
        start.getTime() + 86400000
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

    const results = [];

    let onTime = 0;
    let late = 0;
    let noCheckIn = 0;
    let unregistered = 0;

    // Registered employees.
    for (const [userId, employee] of roster) {
        if (!employee.active) continue;

        const checkIn = checkIns.get(userId);

        let time = "";
        let status = "NO CHECK-IN";
        let messageId = "";

        if (checkIn) {
            time = Utilities.formatDate(
                checkIn.timestamp,
                ATTENDANCE.timezone,
                "HH:mm:ss"
            );

            messageId = checkIn.messageId;
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

    SpreadsheetApp.flush();

    // ========================================
    // 12. EXECUTION SUMMARY
    // ========================================

    console.log("ATTENDANCE COMPLETED");
    console.log("Date: " + date);
    console.log("On time: " + onTime);
    console.log("Late: " + late);
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
    const sheet = ss.getSheetByName("Attendance");

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
        unregistered: 0
    };

    data.forEach(row => {
        const employee = row[2];
        const status = String(row[3]).trim();
        const checkIn = String(row[4]).trim();

        let background = "#FFFFFF";
        let fontColor = "#202124";

        if (status === "NO CHECK-IN") {
            background = "#FEF3C7";
            fontColor = "#92400E";
            summary.missing++;

        } else if (
            status === "PRESENT - NOT ON ROSTER"
        ) {
            background = "#DBEAFE";
            fontColor = "#1E40AF";
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
                        background = "#FEE2E2";
                        fontColor = "#991B1B";
                        summary.late++;

                        console.log(
                            "LATE: " + employee +
                            " | " + checkIn
                        );

                    } else {
                        // GREEN: On time
                        background = "#DCFCE7";
                        fontColor = "#166534";
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

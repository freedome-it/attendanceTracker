const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

const employees = JSON.parse(
    fs.readFileSync(path.join(__dirname, 'employees.json'), 'utf8')
);

const monthsAndDays = {
    gennaio: 31,
    febbraio: 28,
    marzo: 31,
    aprile: 30,
    maggio: 31,
    giugno: 30,
    luglio: 31,
    agosto: 31,
    settembre: 30,
    ottobre: 31,
    novembre: 30,
    dicembre: 31
};

const daysNames = ['L', 'M', 'M', 'G', 'V', 'S', 'D'];

const year = '2026';
const yearStartsOn = 3; // 0 = Monday

// =======================
// STYLES
// =======================

const headerFont = { bold: true };

const centerAlignment = {
    vertical: 'middle',
    horizontal: 'center'
};

const yellowFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFEB3B' }
};

const grayFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFD9D9D9' }
};

const thinBorder = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
};

// =======================
// MAIN
// =======================

async function generateAttendanceXlsx() {
    const workbook = new ExcelJS.Workbook();

    for (const month of Object.keys(monthsAndDays)) {
        const daysInMonth = monthsAndDays[month];
        const sheet = workbook.addWorksheet(`${month}-${year}`);

        for (const employee of employees) {
            // Month + weekday row
            const monthRow = sheet.addRow(
                compileMonthAndDayRow(month, daysInMonth, year, yearStartsOn)
            );

            monthRow.eachCell(cell => {
                cell.font = headerFont;
                cell.alignment = centerAlignment;
                cell.border = thinBorder;
            });

            // Name + day numbers row
            const employeeRow = sheet.addRow(
                compileNameAndDateRow(
                    employee.name,
                    employee.code,
                    monthsAndDays,
                    month
                )
            );

            // Name
            employeeRow.getCell(1).font = headerFont;
            employeeRow.getCell(1).fill = yellowFill;
            employeeRow.getCell(1).border = thinBorder;

            // Code
            employeeRow.getCell(2).font = headerFont;
            employeeRow.getCell(2).fill = yellowFill;
            employeeRow.getCell(2).alignment = centerAlignment;
            employeeRow.getCell(2).border = thinBorder;

            // Day numbers
            employeeRow.eachCell((cell, colNumber) => {
                if (colNumber >= 3) {
                    cell.fill = grayFill;
                    cell.alignment = centerAlignment;
                }
                cell.border = thinBorder;
            });

            // Status rows
            const statusRows = [
                'Lavorato',
                'Smart',
                'Ferie',
                'R.O.L.',
                'Malattia',
                'Chiusura aziendale',
                'Varie'
            ];

            for (const label of statusRows) {
                const row = sheet.addRow([label]);

                row.eachCell(cell => {
                    cell.border = thinBorder;
                    cell.alignment = { vertical: 'middle' };
                });
            }

            // Empty row between employees
            sheet.addRow([]);
        }

        // Column sizing
        sheet.columns.forEach((col, index) => {
            col.width = index < 2 ? 18 : 4;
        });
    }

    const outputPath = path.join(
        __dirname,
        'generated',
        `attendance_${year}.xlsx`
    );

    await workbook.xlsx.writeFile(outputPath);
}

generateAttendanceXlsx();

// =======================
// HELPERS
// =======================

function compileMonthAndDayRow(month, daysInMonth, year, yearStartsOn) {
    const row = [`${month}-${year}`];

    for (let i = 0; i <= daysInMonth; i++) {
        row.push(daysNames[(yearStartsOn + i) % 7]);
    }

    return row;
}

function compileNameAndDateRow(name, code, monthsAndDays, month) {
    const row = [name, code];

    let previousMonthsDays = 0;
    for (const m of Object.keys(monthsAndDays)) {
        if (m === month) break;
        previousMonthsDays += monthsAndDays[m];
    }

    for (let i = 0; i < monthsAndDays[month]; i++) {
        row.push(previousMonthsDays + i + 1);
    }

    return row;
}

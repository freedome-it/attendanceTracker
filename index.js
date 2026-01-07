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
const yearStartsOn = 3; // 0 = monday

const companyClosures = {
    gennaio: [1, 2, 5, 6],
    febbraio: [],
    marzo: [],
    aprile: [3, 5, 6, 25],
    maggio: [1],
    giugno: [2],
    luglio: [],
    agosto: [15, 16],
    settembre: [],
    ottobre: [],
    novembre: [1],
    dicembre: [7, 8, 25, 26, 28, 29, 30, 31]
}

let SWDaysStart = {
    first: 3,
    second:4,
    dir: 'bw' // 'bw' = backward, 'fw' = forward
}

let SWDaysUpdate = { ...SWDaysStart };

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

// =======================
// MAIN
// =======================

async function generateAttendanceXlsx() {
    for (const team in employees) {
        const workbook = new ExcelJS.Workbook();
        let daysFromStartOfYear = yearStartsOn;
    
        for (const month of Object.keys(monthsAndDays)) {
            const daysInMonth = monthsAndDays[month];
            const sheet = workbook.addWorksheet(`${month} ${year}`);
    
            for (const employee of employees[team]) {
                // month + weekday row
                const monthRow = sheet.addRow(
                    compileMonthAndDayRow(month, daysInMonth, year, daysFromStartOfYear)
                );
    
                monthRow.eachCell(cell => {
                    cell.font = headerFont;
                    cell.alignment = centerAlignment;
                });
    
                // name + day numbers row
                const employeeRow = sheet.addRow(
                    compileNameAndDateRow(
                        employee.name,
                        employee.code,
                        monthsAndDays,
                        month
                    )
                );
    
                // name
                employeeRow.getCell(1).font = headerFont;
                employeeRow.getCell(1).fill = yellowFill;
    
                // code
                employeeRow.getCell(2).font = headerFont;
                employeeRow.getCell(2).fill = yellowFill;
                employeeRow.getCell(2).alignment = centerAlignment;
    
                // day numbers
                employeeRow.eachCell((cell, colNumber) => {
                    if (colNumber >= 3) {
                        cell.fill = grayFill;
                        cell.alignment = centerAlignment;
                    }
                });
    
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
                    const row = label === 'Smart'
                        ? sheet.addRow(compileSmartWorkingRow(month, daysInMonth, daysFromStartOfYear))
                        : label === 'Chiusura aziendale' 
                            ? sheet.addRow(compileCompanyClosureRow(month, daysInMonth))
                            : sheet.addRow([label]);
    
                    row.eachCell(cell => {
                        cell.alignment = { vertical: 'middle'};
                    });
                }
    
                // empty row between employees
                sheet.addRow([]);
            }
    
            // column sizing
            sheet.columns.forEach((col, index) => {
                col.width = index < 1 ? 24 : 6;
            });
            
            SWDaysStart = { ...SWDaysUpdate };
            daysFromStartOfYear += daysInMonth;
        }
    
        const outputPath = path.join(
            __dirname,
            'generated',
            team,
            `attendance_${year}.xlsx`
        );

        // Create directory structure if it doesn't exist
        const outputDir = path.dirname(outputPath);
        fs.mkdirSync(outputDir, { recursive: true });
        
        await workbook.xlsx.writeFile(outputPath);

        // Reset SWDaysStart for next team
        SWDaysStart = {
            first: 3,
            second:4,
            dir: 'bw' // 'bw' = backward, 'fw' = forward
        }

        SWDaysUpdate = { ...SWDaysStart };
    }
}

generateAttendanceXlsx();

// =======================
// HELPERS
// =======================

function compileMonthAndDayRow(month, daysInMonth, year, daysFromStartOfYear) {
    const row = [`${month}-${year}`];

    row.push('');

    for (let i = 0; i < daysInMonth; i++) {
        row.push(daysNames[(daysFromStartOfYear + i) % 7]);
    }

    return row;
}

function compileNameAndDateRow(name, code, monthsAndDays, month) {
    const row = [name, code];

    for (let i = 0; i < monthsAndDays[month]; i++) {
        row.push(i + 1);
    }

    return row;
}

function compileSmartWorkingRow(month, daysInMonth, daysFromStartOfYear) {
    const row = ['Smart'];
    
    row.push('');

    let SWDays = { ...SWDaysStart };
    
    for (let i = 0; i < daysInMonth; i++) {
        const currentDayOfWeek = (daysFromStartOfYear + i) % 7;


        if (currentDayOfWeek === SWDays.first || currentDayOfWeek === SWDays.second) {
            if(companyClosures[month].includes(i + 1)){
                row.push('');
            }else{
                row.push('X');
            }
        } else {
            row.push('');
        }

        // Update SWDays with 'Ping Pong' logic
        if (currentDayOfWeek === 6) { // Sunday
            if (SWDays.dir === 'bw') {
                if(SWDays.first === 0){
                    SWDays.first++;
                    SWDays.second++;
                    SWDays.dir = 'fw';
                }else{
                    SWDays.first--;
                    SWDays.second--;
                }
            }
            else if (SWDays.dir === 'fw') {
                if(SWDays.first === 3){
                    SWDays.first--;
                    SWDays.second--;
                    SWDays.dir = 'bw';
                }else{
                    SWDays.first++;
                    SWDays.second++;
                }
            }
        }
    }
    SWDaysUpdate = {...SWDays}
    return row;
}

function compileCompanyClosureRow(month, daysInMonth) {
    const row = ['Chiusura aziendale'];

    row.push('');
    
    for (let i = 0; i < daysInMonth; i++) {
        if (companyClosures[month].includes(i + 1)) {
            row.push('X');
        } else {
            row.push('');
        }
    }

    return row;
}

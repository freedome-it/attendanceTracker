fs = require('fs');
const path = require('path');
const employees = JSON.parse(fs.readFileSync(path.join(__dirname, 'employees.json'), 'utf8'));

//foreach employee, write a csv this way:

//gennaio-26	
//Mario Rossi	E001
//Lavorato	
//Smart	
//Ferie	
//R.O.L.	
//Malattia	
//Chiusura aziendale	
//Varie	

const monthsAndDays = {
    'gennaio': 31,
    'febbraio': 28,
    'marzo': 31,
    'aprile': 30,
    'maggio': 31,
    'giugno': 30,
    'luglio': 31,
    'agosto': 31,
    'settembre': 30,
    'ottobre': 31,
    'novembre': 30,
    'dicembre': 31
};

const daysNames = [
    'L',
    'M',
    'M',
    'G',
    'V',
    'S',
    'D'
]

year = '2026';
yearsStartsOn = 3

for (let month in monthsAndDays) {
    let monthCSVContent = '';
    for (let employee of employees) {
        monthCSVContent += compileMonthAndDayLine(month, monthsAndDays[month], year, yearsStartsOn) + '\n';
        monthCSVContent += compileNameAndDateLine(employee.name, employee.code, monthsAndDays, month) + '\n';
        monthCSVContent += 'Lavorato' + '\n';
        monthCSVContent += 'Smart' + '\n';
        monthCSVContent += 'Ferie' + '\n';
        monthCSVContent += 'R.O.L.' + '\n';
        monthCSVContent += 'Malattia' + '\n';
        monthCSVContent += 'Chiusura aziendale' + '\n';
        monthCSVContent += 'Varie' + '\n';
    }
    fs.writeFileSync(path.join(__dirname + '/generated', `attendance_${month}_${year}.csv`), monthCSVContent);
}

function compileMonthAndDayLine(month, daysInMonth, year, yearsStartsOn) {
    let line = `${month}-${year}\t`;
    for (let i = 0; i < daysInMonth; i++) {
        line += daysNames[(yearsStartsOn + i) % 7] + '\t';
    }
    return line.trim();
}

function compileNameAndDateLine(name, code, monthsAndDays, month) {
    let line = `${name}\t${code}`;
    let previousMonthsDays = 0;
    for (let m in monthsAndDays) {
        if (m === month) {
            break;
        }
        previousMonthsDays += monthsAndDays[m];
    }
    for (let i = 0; i < monthsAndDays[month]; i++) {
        line += `\t${previousMonthsDays + i + 1}`;
    }
    return line;
}


const fs = require('fs');
const XLSX = require('xlsx');

// Function to convert numeric values to human-readable date-time format
function convertNumericToDateTime(numericValue) {
    const epochDate = new Date('1900-01-01');
    const date = new Date(epochDate);
    date.setDate(epochDate.getDate() + Math.floor(numericValue));
    date.setMilliseconds((numericValue % 1) * 24 * 60 * 60 * 1000);
    return date.toLocaleString('en-US', { timeZone: 'UTC' });
}

// Function to convert numeric hours to time duration string
function convertNumericHoursToTime(numericHours) {
    const hours = Math.floor(numericHours);
    const minutes = Math.round((numericHours % 1) * 60);
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
}

// Function to convert time duration string to numeric hours
function convertTimeToNumericHours(timeString) {
    const [hours, minutes] = timeString.split(':').map(Number);
    return hours + minutes / 60;
}


// Function to find and print employee names for 7 consecutive dates
function findEmployeesForConsecutiveDates(file_path, consecutiveDays) {
    try {
        const fileContent = fs.readFileSync(file_path);
        const workbook = XLSX.read(fileContent, { type: 'buffer' });

        const sheet_name = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheet_name];

        const dateColumnName = 'Time';
        const employeeNameColumnName = 'Employee Name';

        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        const dateIndex = rows[0].indexOf(dateColumnName);
        const employeeNameIndex = rows[0].indexOf(employeeNameColumnName);

        let employeeDatesMap = new Map(); // Map to store dates for each employee

        for (let i = 1; i < rows.length; i++) {
            const dateCellValue = rows[i][dateIndex];
            const employeeNameCellValue = rows[i][employeeNameIndex];

            if (dateCellValue !== undefined && employeeNameCellValue !== undefined) {
                const formattedDateTime = convertNumericToDateTime(dateCellValue);
                const formattedDateWithoutTime = formattedDateTime.replace(/, \d+:\d+:\d+ [APMapm]+$/, '');

                if (!employeeDatesMap.has(employeeNameCellValue)) {
                    employeeDatesMap.set(employeeNameCellValue, [formattedDateWithoutTime]);
                } else {
                    const dates = employeeDatesMap.get(employeeNameCellValue);

                    if (dates.includes(formattedDateWithoutTime)) {
                        // Skip duplicate dates for the same employee
                        continue;
                    }

                    dates.push(formattedDateWithoutTime);
                    dates.sort(); // Sort dates to identify consecutive ones

                    if (checkConsecutiveDates(dates, consecutiveDays)) {
                        fs.appendFileSync('output.txt', `\nTask 3: Employee Name: ${employeeNameCellValue}\n`);
                        fs.appendFileSync('output.txt', `Consecutive Dates: ${dates.join(', ')}\n`);
                        break; 
                        // Exit the loop once 7 consecutive dates are found
                    }
                }
            }
        }
    } catch (error) {
        console.error(`Error: ${error.message}`);
    }
}

// Function to check if dates are consecutive
function checkConsecutiveDates(dates, consecutiveDays) {
    for (let i = 0; i < dates.length - consecutiveDays + 1; i++) {
        const current = new Date(dates[i]);
        const next = new Date(dates[i + consecutiveDays - 1]);
        const diffTime = Math.abs(next - current);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays === consecutiveDays - 1) {
            return true; // Found consecutive dates
        }
    }

    return false; // Consecutive dates not found
}

// Function to analyze Excel file for tasks 1 and 2
function analyzeExcelFile(file_path) {
    try {
        // Read the Excel file
        const workbook = XLSX.readFile(file_path);
        const sheet_name = workbook.SheetNames[0]; // Assuming data is in the first sheet
        const sheet = workbook.Sheets[sheet_name];

        // Convert sheet data to JSON
        const data = XLSX.utils.sheet_to_json(sheet);

        // Adjust column names based on your Excel file
        const columnNameMap = {
            positionStatus: 'Position Status',
            time: 'Time',
            timeOut: 'Time Out',
            hoursWorked: 'Timecard Hours (as Time)',
            payCycleEndDate: 'Pay Cycle End Date',
            employeeName: 'Employee Name',
            fileNumber: 'File Number'
        };

        // Task 1: Identify employees who have worked more than 1 hour or less than 10 hours
        const filteredEmployeesTask1 = data.filter(emp => {
            try {
                const hoursWorked = convertTimeToNumericHours(emp[columnNameMap.hoursWorked]);
                if (isNaN(hoursWorked)) {
                    console.log(`Invalid value in hoursWorked for ${emp[columnNameMap.employeeName]}: ${emp[columnNameMap.hoursWorked]}`);
                }
                return hoursWorked > 1 && hoursWorked < 10; // More than 1 hour and less than 10 hours
            } catch (error) {
                // Log and continue with the next employee if there's an error
                console.log(`Error processing hoursWorked for ${emp[columnNameMap.employeeName]}: ${error.message}`);
                return false;
            }
        });

        // Log the results for Task 1 to the console and file
        fs.writeFileSync('output.txt', "Task 1: Employees who have worked more than 1 hour and less than 10 hours:\n");

        if (filteredEmployeesTask1.length > 0) {
            filteredEmployeesTask1.forEach(emp => {
                const logData = `Employee Name: ${emp[columnNameMap.employeeName]}\nFile Number: ${emp[columnNameMap.fileNumber]}\nHours Worked: ${emp[columnNameMap.hoursWorked]}\nFormatted Hours Worked: ${convertTimeToNumericHours(emp[columnNameMap.hoursWorked])}\n`;
                console.log(logData);
                fs.appendFileSync('output.txt', logData + '\n');
            });
        } else {
            console.log("No employees found who have worked more than 1 hour and less than 10 hours.");
            fs.appendFileSync('output.txt', "No employees found who have worked more than 1 hour and less than 10 hours.\n");
        }

        // Task 2: Identify employees who have worked for more than 14 hours in a single shift
        const longShiftEmployees = data.filter(emp => {
            try {
                const hoursWorked = parseFloat(emp[columnNameMap.hoursWorked]);
                return hoursWorked > 14;
            } catch (error) {
                // Log and continue with the next employee if there's an error
                return false;
            }
        });

        // Log the results for Task 2 to the console and file
        fs.appendFileSync('output.txt', "\nTask 2: Employees who have worked for more than 14 hours in a single shift:\n");

        if (longShiftEmployees.length > 0) {
            longShiftEmployees.forEach(emp => {
                const logData = `Employee Name: ${emp[columnNameMap.employeeName]}\nFile Number: ${emp[columnNameMap.fileNumber]}\nHours Worked: ${emp[columnNameMap.hoursWorked]}\nFormatted Hours Worked: ${convertNumericHoursToTime(emp[columnNameMap.hoursWorked])}\n`;
                console.log(logData);
                fs.appendFileSync('output.txt', logData + '\n');
            });
        } else {
            console.log("No employees found who have worked for more than 14 hours in a single shift.");
            fs.appendFileSync('output.txt', "No employees found who have worked for more than 14 hours in a single shift.\n");
        }

        // Task 3: Find and print employee names for 7 consecutive dates
        findEmployeesForConsecutiveDates(file_path, 7);

    } catch (error) {
        console.error(`Error: ${error.message}`);
    }
}

// Call the function with the path to the Excel file
analyzeExcelFile('assets/Assignment_Timecard2.xlsx');

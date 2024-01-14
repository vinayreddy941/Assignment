const fs = require('fs');
const XLSX = require('xlsx');

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

        // Task 1: Employees who have worked for 7 consecutive days
        const consecutiveDaysEmployees = data.filter((emp, index, array) => {
            return index > 0 && (new Date(emp[columnNameMap.payCycleEndDate]) - new Date(array[index - 1][columnNameMap.payCycleEndDate])) / (24 * 60 * 60 * 1000) === 1;
        });

        // Task 2a: Employees who have less than 10 hours between shifts but greater than 1 hour
        const timeBetweenShiftsEmployees = data.filter((emp, index, array) => {
            return index > 0 && (new Date(emp[columnNameMap.time]) - new Date(array[index - 1][columnNameMap.timeOut])) / (60 * 60 * 1000) < 10 &&
                (new Date(emp[columnNameMap.time]) - new Date(array[index - 1][columnNameMap.timeOut])) / (60 * 60 * 1000) > 1;
        });

        // Task 2b: Employees who have worked for more than 14 hours in a single shift
        const longShiftEmployees = data.filter(emp => parseFloat(emp[columnNameMap.hoursWorked]) > 14);

        
    } catch (error) {
        console.error(`Error: ${error.message}`);
    }
}

analyzeExcelFile('assets/Assignment_Timecard.xlsx');

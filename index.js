const xlsx = require('xlsx');
const fs = require('fs');
const db = require('./db');
require('dotenv').config();

// Function to read Excel file and convert to CSV
const excelToCsv = (excelFilePath, csvFilePath) => {
    const workbook = xlsx.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const csvData = xlsx.utils.sheet_to_csv(worksheet);
    fs.writeFileSync(csvFilePath, csvData);

    return { csvData, sheetName };
};

const createTableFromCsv = async (csvData, sheetName) => {
    const lines = csvData.split('\n');
    const headers = lines[0].split(',');

    // Create table SQL with dynamically generated table name based on sheet name
    const createTableQuery = `
        CREATE TABLE IF NOT EXISTS ${sheetName} (
            ${headers.map(header => `${header.trim()} TEXT`).join(', ')}
        );
    `;

    try {
        const result = await db.query(createTableQuery);
        console.log(result.command)
        if (result.command === 'CREATE') {
            console.log(`New table "${sheetName}" created successfully.`);
        } else if (result.command === 'SELECT') {
            console.log(`Table "${sheetName}" already exists.`);
        } else {
            console.log(`Unexpected result: ${result.command}`);
        }
    } catch (err) {
        console.error('Error creating or checking table:', err);
    }
};

// Function to insert CSV data into table based on sheet name
const insertDataFromCsv = async (csvFilePath, sheetName) => {
    const csvData = fs.readFileSync(csvFilePath, 'utf-8');
    const lines = csvData.split('\n').slice(1); // Skip header line
    const headers = csvData.split('\n')[0].split(',');

    for (const line of lines) {
        if (line.trim()) {
            const values = line.split(',').map(value => value.trim());
            const insertQuery = `
        INSERT INTO ${sheetName} (${headers.join(', ')})
        VALUES (${values.map(value => `'${value}'`).join(', ')});
      `;

            try {
                await db.query(insertQuery);
                console.log(`Data inserted successfully into table "${sheetName}".`);
            } catch (err) {
                console.error(`Error inserting data into table "${sheetName}":`, err);
            }
        }
    }
};

// Main function
(async () => {
    const excelFilePath = './Book1.xlsx';
    const csvFilePath = './employee.csv';

    // Step 1: Convert Excel to CSV
    const { csvData, sheetName } = excelToCsv(excelFilePath, csvFilePath);

    // Step 2: Create table based on CSV data
    await createTableFromCsv(csvData, sheetName);

    // Step 3: Insert data from CSV into table
    await insertDataFromCsv(csvFilePath, sheetName);
})();

const XLSX = require("xlsx");
const fs = require("fs");

// Load the original Excel file
const originalFilePath = "master_list_full_formatted.xlsx";
const workbook = XLSX.readFile(originalFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert worksheet data to an array of objects
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
const headerRow = data[0];
const dataRows = data.slice(1);

const default_data = "default";
const emailColumnIndex = headerRow.indexOf("email");
const phoneColumnIndex = headerRow.indexOf("phone");
const nameColumnIndex = headerRow.indexOf("name");
const addressColumnIndex = headerRow.indexOf("address");
const stateColumnIndex = headerRow.indexOf("state");
const zipcodeColumnIndex = headerRow.indexOf("zipcode");

// Fill empty "email" cells with a default value
for (let i = 0; i < dataRows.length; i++) {
  if (!dataRows[i][emailColumnIndex]) {
    dataRows[i][emailColumnIndex] = "default@example.com"; // Replace with your default email
  }
  if (!dataRows[i][phoneColumnIndex]) {
    dataRows[i][phoneColumnIndex] = "(333) 333-3333"; // Replace with your default phone
  }
  if (!dataRows[i][nameColumnIndex]) {
    dataRows[i][nameColumnIndex] = default_data; // Replace with your default phone
  }
  if (!dataRows[i][addressColumnIndex]) {
    dataRows[i][addressColumnIndex] = default_data; // Replace with your default phone
  }
  if (!dataRows[i][stateColumnIndex]) {
    dataRows[i][stateColumnIndex] = default_data; // Replace with your default phone
  }
  if (!dataRows[i][zipcodeColumnIndex]) {
    dataRows[i][zipcodeColumnIndex] = "00000"; // Replace with your default phone
  }
}

// Calculate the total number of rows and the number of files needed
const maxRowsPerFile = 900;
const numFiles = Math.ceil(dataRows.length / maxRowsPerFile);

// Divide the data into smaller arrays and save as separate files
for (let i = 0; i < numFiles; i++) {
  const startRow = i * maxRowsPerFile;
  const endRow = Math.min((i + 1) * maxRowsPerFile, dataRows.length);
  const subsetData = dataRows.slice(startRow, endRow);

  const subsetWorkbook = XLSX.utils.book_new();
  const subsetWorksheet = XLSX.utils.aoa_to_sheet([headerRow, ...subsetData]);
  XLSX.utils.book_append_sheet(subsetWorkbook, subsetWorksheet, sheetName);

  const newFilePath = `output_file_${i + 1}.xlsx`;
  XLSX.writeFile(subsetWorkbook, newFilePath);

  console.log(`${newFilePath} created.`);
}

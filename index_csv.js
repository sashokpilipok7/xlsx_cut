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

// Define column indices to remove (replace with actual indices)
const columnsToRemove = 24; // Example: Remove the second and fourth columns

// Remove specified columns from the data
const modifiedData = dataRows.map((row) =>
  row.filter((_, index) => index < columnsToRemove)
);

// Calculate the total number of rows and the number of files needed
const maxRowsPerFile = 2000;
const numFiles = Math.ceil(modifiedData.length / maxRowsPerFile);

// Divide the data into smaller arrays and save as separate CSV files
for (let i = 0; i < numFiles; i++) {
  const startRow = i * maxRowsPerFile;
  const endRow = Math.min((i + 1) * maxRowsPerFile, modifiedData.length);
  const subsetData = modifiedData.slice(startRow, endRow);

  const csvContent = [headerRow.join(",")]
    .concat(subsetData.map((row) => row.join(",")))
    .join("\n");

  const newFilePath = `output_file_${i + 1}.csv`;
  fs.writeFileSync(newFilePath, csvContent);

  console.log(`${newFilePath} created.`);
}

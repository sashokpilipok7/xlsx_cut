const XLSX = require("xlsx");
const fs = require("fs");

// Replace 'input.xlsx' with the name of your original XLSX file
const workbook = XLSX.readFile("master_list_full_formatted.xlsx");
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];

const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
const columnNames = rows[0]; // Assuming the first row contains column names
const first10Rows = rows.slice(1, 11); // Skip the header row and get first 10 data rows

const default_data = "default";
const emailColumnIndex = columnNames.indexOf("email");
const phoneColumnIndex = columnNames.indexOf("phone");
const nameColumnIndex = columnNames.indexOf("name");
const addressColumnIndex = columnNames.indexOf("address");
const stateColumnIndex = columnNames.indexOf("state");
const zipcodeColumnIndex = columnNames.indexOf("zipcode");

// Fill empty "email" cells with a default value
for (let i = 0; i < first10Rows.length; i++) {
  if (!first10Rows[i][emailColumnIndex]) {
    first10Rows[i][emailColumnIndex] = "default@example.com"; // Replace with your default email
  }
  if (!first10Rows[i][phoneColumnIndex]) {
    first10Rows[i][phoneColumnIndex] = "(333) 333-3333"; // Replace with your default phone
  }
  if (!first10Rows[i][nameColumnIndex]) {
    first10Rows[i][nameColumnIndex] = default_data; // Replace with your default phone
  }
  if (!first10Rows[i][addressColumnIndex]) {
    first10Rows[i][addressColumnIndex] = default_data; // Replace with your default phone
  }
  if (!first10Rows[i][stateColumnIndex]) {
    first10Rows[i][stateColumnIndex] = default_data; // Replace with your default phone
  }
  if (!first10Rows[i][zipcodeColumnIndex]) {
    first10Rows[i][zipcodeColumnIndex] = "00000"; // Replace with your default phone
  }
}

// Add a new column "email" with the value "test@test.com" to each row
// const rowsWithEmail = first10Rows.map((row) => [...row, "test@test.com"]);

// // Insert "https:" at the start of the "Top_Image_URL" column in each row
// const rowsWithModifiedURL = rowsWithEmail.map((row) => {
//   const [topImageURL, ...restColumns] = row;
//   const modifiedTopImageURL = `https:${topImageURL}`;
//   return [modifiedTopImageURL, ...restColumns];
// });

// Create a new array with the updated rows (including column names)
const updatedData = [columnNames, ...first10Rows];

// Create a new workbook with the updated rows
const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.aoa_to_sheet(updatedData);
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

// Replace 'output.xlsx' with the desired name of the new XLSX file
XLSX.writeFile(newWorkbook, "output.xlsx");

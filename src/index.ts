// Entry point for Excel data merging CLI tool
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

const dataDir = path.join(__dirname, '../data');
const files = fs.readdirSync(dataDir);
console.log('Files in data folder:');
files.forEach(file => console.log(file));

// Filter for only .xlsx files (excluding .gitkeep and other non-Excel files)
const excelFiles = files.filter(file => file.endsWith('.xlsx'));

// Read and merge Excel files
let mergedData: any[] = [];
excelFiles.forEach(file => {
  const filePath = path.join(dataDir, file);
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(sheet);
  mergedData = mergedData.concat(json);
});

// console.log('Merged Data:');
// console.log(mergedData);

// // Optionally, write merged data to a new Excel file
// const outputWb = XLSX.utils.book_new();
// const outputWs = XLSX.utils.json_to_sheet(mergedData);
// XLSX.utils.book_append_sheet(outputWb, outputWs, 'Merged');
// XLSX.writeFile(outputWb, path.join(dataDir, 'merged_output.xlsx'));
// console.log('Merged Excel file written to data/merged_output.xlsx');

// console.log('Excel Data Merger CLI - To be implemented');
// TODO: Implement CLI argument parsing, merging logic, and export functionality

// // List header fields for each Excel file
// excelFiles.forEach(file => {
//   const filePath = path.join(dataDir, file);
//   const workbook = XLSX.readFile(filePath);
//   const sheetName = workbook.SheetNames[0];
//   const sheet = workbook.Sheets[sheetName];
//   const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
//   const headers = json[0];
//   console.log(`Headers in ${file}:`, headers);
// });

// Only merge the two specified files
const contactFile = 'Contact Information (Responses).xlsx';
const diasporaFile = 'TARM-Updated-Diaspora-Disciples.xlsx';

let contactData: any[] = [];
let diasporaData: any[] = [];

// Read Contact Information (Responses).xlsx
if (files.includes(contactFile)) {
  const filePath = path.join(dataDir, contactFile);
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  contactData = XLSX.utils.sheet_to_json(sheet);
}
console.log('Contact Data:');
console.log(contactData.slice(0, 5)); // Log first 5 rows for preview
// console.log(contactData.length, 'rows read from Contact Information (Responses).xlsx');
// Read TARM-Updated-Diaspora-Disciples.xlsx
if (files.includes(diasporaFile)) {
  const filePath = path.join(dataDir, diasporaFile);
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  diasporaData = XLSX.utils.sheet_to_json(sheet);
}

// Build a map for fast lookup
function getKeyFromContactInfo(row: any) {
  if (row['Email Address']) {
    return String(row['Email Address']).trim().toLowerCase();
  }
  return `${row['First Name'] || ''} ${row['Last name'] || ''}`.trim().toLowerCase();
}
function getKeyFromDiaspora(row: any) {
  if (row['Email']) {
    return String(row['Email']).trim().toLowerCase();
  }
  return `${row['NAME'] || ''} ${row['SURNAME'] || ''}`.trim().toLowerCase();
}

const diasporaMap = new Map<string, any>();
diasporaData.forEach(row => {
  diasporaMap.set(getKeyFromDiaspora(row), row);
});

// Merge rows by unique key
const merged: any[] = [];
contactData.forEach(row => {
  const key = getKeyFromContactInfo(row);
  const diasporaRow = diasporaMap.get(key);
  if (diasporaRow) {
    merged.push({ ...row, ...diasporaRow });
  } else {
    merged.push(row);
  }
});

diasporaData.forEach(row => {
  const key = getKeyFromDiaspora(row);
  if (!contactData.some(r => getKeyFromContactInfo(r) === key)) {
    merged.push(row);
  }
});

console.log('Merged Data:');
// console.log(merged);

// Convert Excel serial date to string in 'M/D/YYYY H:mm:ss' format
function excelSerialToDateString(serial: number): string {
  // Excel's epoch starts at 1899-12-30
  const excelEpoch = new Date(Date.UTC(1899, 11, 30));
  const msPerDay = 24 * 60 * 60 * 1000;
  const days = Math.floor(serial);
  const ms = Math.round((serial - days) * msPerDay);
  const date = new Date(excelEpoch.getTime() + days * msPerDay + ms);
  // Format as 'M/D/YYYY H:mm:ss'
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()} ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

// Remove duplicate NAME, SURNAME fields from merged output
const mergedStringDates = merged.map(row => {
  const newRow: any = {};
  for (const key in row) {
    if (row.hasOwnProperty(key)) {
      // Merge both Wedding Anniversary fields into a single 'Wedding Anniversary' column
      if (key === 'When is your marriage anniversary?' || key === 'Wedding Anniversary') {
        if (!('Wedding Anniversary' in newRow)) {
          newRow['Wedding Anniversary'] = row['Wedding Anniversary'] || row['When is your marriage anniversary?'];
        }
        continue;
      }
      // Merge both Phone Number fields into a single 'Phone Number (Please add country code e.g. +44 for UK lines)' column
      if (key === 'Phone Number (Please add country code e.g. +44 for UK lines)' || key === 'Phone number') {
        if (!('Phone Number (Please add country code e.g. +44 for UK lines)' in newRow)) {
          newRow['Phone Number (Please add country code e.g. +44 for UK lines)'] = row['Phone Number (Please add country code e.g. +44 for UK lines)'] || row['Phone number'];
        }
        continue;
      }
      // Merge both Email fields into a single 'Email Address' column
      if (key === 'Email Address' || key === 'Email') {
        if (!('Email Address' in newRow)) {
          newRow['Email Address'] = row['Email Address'] || row['Email'];
        }
        continue;
      }
      // Merge both Date of Birth fields into a single 'Date of Birth' column
      if (key === 'Date of Birth' || key === 'DOB') {
        if (!('Date of Birth' in newRow)) {
          let dobVal = row['Date of Birth'] || row['DOB'];
          // Convert Excel serial number to date string if needed
          if (typeof dobVal === 'number' && dobVal > 20000 && dobVal < 90000) {
            dobVal = excelSerialToDateString(dobVal).split(' ')[0]; // Only keep the date part
          }
          newRow['Date of Birth'] = dobVal;
        }
        continue;
      }
      // Merge both Marrital Status fields into a single 'Marrital Status' column
      if (key === 'Marrital Status' || key === 'Are you married?') {
        if (!('Marrital Status' in newRow)) {
          newRow['Marrital Status'] = row['Marrital Status'] || row['Are you married?'];
        }
        continue;
      }
      // Merge both country fields into a single 'Country' column
      if (key === 'Country of Residence' || key === 'Please state the country where you currently reside?') {
        if (!('Country' in newRow)) {
          newRow['Country'] = row['Country of Residence'] || row['Please state the country where you currently reside?'];
        }
        continue;
      }
      // Skip duplicate NAME and SURNAME if First Name and Last name exist
      if ((key === 'NAME' && 'First Name' in row) || (key === 'SURNAME' && 'Last name' in row)) {
        continue;
      }
      let val = row[key];
      // Convert Excel serial numbers to date string if in valid range
      if (typeof val === 'number' && val > 20000 && val < 90000) {
        val = excelSerialToDateString(val);
      }
      newRow[key] = val;
    }
  }
  return newRow;
});

// Write merged data to a new Excel file (Desktop only)
const outputWb = XLSX.utils.book_new();
const outputWs = XLSX.utils.json_to_sheet(mergedStringDates, { cellDates: false });
XLSX.utils.book_append_sheet(outputWb, outputWs, 'Merged');
const desktopOutputPath = 'C:/Users/Micheal/OneDrive - Swansea University/Desktop/merged_output.xlsx';
try {
  XLSX.writeFile(outputWb, desktopOutputPath, { cellDates: false });
  console.log('Merged Excel file written to Desktop as merged_output.xlsx');
} catch (err) {
  if (err && typeof err === 'object' && 'code' in err && (err as any).code === 'EBUSY') {
    console.error('Error: merged_output.xlsx is open on your Desktop. Please close it and try again.');
  } else {
    console.error('Error writing merged_output.xlsx to Desktop:', err);
  }
}

// Entry point for Excel data merging CLI tool
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

const dataDir = path.join(__dirname, '../data');
const files = fs.readdirSync(dataDir);
console.log('Files in data folder:');
files.forEach(file => console.log(file));

console.log('Excel Data Merger CLI - To be implemented');
// TODO: Implement CLI argument parsing, merging logic, and export functionality

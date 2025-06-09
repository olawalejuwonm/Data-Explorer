// filepath: src/processMergedData.ts
// Utility functions to process merged Excel data for the CLI tool

import * as XLSX from 'xlsx';

// open the const desktopOutputPath = 'C:/Users/Micheal/OneDrive - Swansea University/Desktop/merged_output.xlsx';

const desktopOutputPath = 'C:/Users/Micheal/OneDrive - Swansea University/Desktop/merged_output.xlsx';


// Define checked-in months columns
export const checkedInMonths = [
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
];

// Remove checked-in months columns from main sheet data
export function removeCheckedInMonthsColumns(data: any[]): any[] {
    return data.map(row => {
        const newRow = { ...row };
        checkedInMonths.forEach(month => delete newRow[month]);
        return newRow;
    });
}

// Prepare checked-in months sheet data
export function extractCheckedInMonthsSheet(data: any[]): any[] {
    return data.map(row => {
        const checkedInStatus = checkedInMonths.reduce((acc, month) => {
            let val = row[month] || '';
            // Normalize to 'Called', 'Not Called', or 'Unreachable'
            if (typeof val === 'string') {
                const lower = val.trim().toLowerCase();
                if (['called', 'yes', 'y', '✓', '✔', 'done', 'called in'].includes(lower)) {
                    val = 'Called';
                } else if (['unreachable', 'not reachable', 'nr', 'n/r'].includes(lower)) {
                    val = 'Unreachable';
                } else if (['not called', 'no', 'n', ''].includes(lower)) {
                    val = 'Not Called';
                } else {
                    val = 'Not Called'; // Default: always Not Called if not matched
                }
            }
            acc[month] = val;
            return acc;
        }, {} as Record<string, any>);
        return {
            'Email Address': row['Email Address'] || '',
            'First Name': row['First Name'] || '',
            'Last name': row['Last name'] || '',
            ...checkedInStatus
        };
    });
}

// Get all unique keys from the merged data
export function getAllKeys(data: any[]): string[] {
    return Array.from(new Set(data.flatMap(row => Object.keys(row))));
}

// Get headers for main sheet: ordered fields, then checked-in months, then the rest
export function getMainSheetHeaders(data: any[]): string[] {
    const orderedHeaders = [
        'First Name',
        'Last name',
        'Phone Number (Please add country code e.g. +44 for UK lines)',
        'Email Address',
        'Country',
        'Date of Birth'
    ];
    // Insert checkedInMonths after 'Date of Birth'
    const idx = orderedHeaders.indexOf('Date of Birth');
    const headersWithMonths = [
        ...orderedHeaders.slice(0, idx + 1),
        ...checkedInMonths,
        ...orderedHeaders.slice(idx + 1)
    ];
    // Add any remaining columns not already included
    const allKeys = getAllKeys(data);
    const rest = allKeys.filter(k => !headersWithMonths.includes(k) && !checkedInMonths.includes(k));
    return [...headersWithMonths, ...rest];
}

// Function to read, process, and write a new Excel file based on desktopOutputPath
export function processAndWriteNewFile(newFilePath: string) {
    // Read the workbook from desktopOutputPath
    const workbook = XLSX.readFile(desktopOutputPath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);

    // Prepare checked-in months sheet data (normalized values)
    const checkedInSheetData = extractCheckedInMonthsSheet(data);
    // Get headers for main sheet
    const mainHeaders = getMainSheetHeaders(data);

    // Merge checked-in months data into main sheet rows
    const mergedRows = data.map((row, idx) => {
        const checked = checkedInSheetData[idx] || {};
        // Defensive: ensure row is an object
        const newRow: Record<string, any> = Object.assign({}, row);
        checkedInMonths.forEach(month => {
            // Only allow the three valid options
            let value = checked[month];
            if (value !== 'Called' && value !== 'Unreachable') {
                value = 'Not Called';
            }
            newRow[month] = value;
        });
        return newRow;
    });

    // Create a new workbook
    const newWb = XLSX.utils.book_new();
    // Add main sheet with checked-in months columns (normalized)
    const mainWs = XLSX.utils.json_to_sheet(mergedRows, { header: mainHeaders });

    // Ensure checked-in months columns are present in every row
    const range = XLSX.utils.decode_range(mainWs['!ref']!);
    for (let R = 0; R <= range.e.r; ++R) { // include header row
        for (let C = 0; C < mainHeaders.length; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!mainWs[cellAddress]) continue;
            // Lock all fields (cells)
            if (!mainWs[cellAddress].s) mainWs[cellAddress].s = {};
            mainWs[cellAddress].s.protection = { locked: true };
        }
    }

    XLSX.utils.book_append_sheet(newWb, mainWs, 'Merged');

    // Write the new workbook to the specified file path
    XLSX.writeFile(newWb, newFilePath);
    console.log(`Processed data written to ${newFilePath}`);
}

processAndWriteNewFile('C:/Users/Micheal/OneDrive - Swansea University/Desktop/processed_output.xlsx');

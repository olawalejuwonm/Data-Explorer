# Excel Data Merger (TypeScript)

This project is a TypeScript CLI tool to merge data from multiple Excel files based on a common key.

## Features
- Read multiple Excel files
- Merge data based on a specified key
- Export merged data to a new Excel file

## Usage
1. Place your Excel files in the project directory.
2. Run the CLI tool with the required arguments (to be implemented).

## Setup
- Install dependencies: `npm install`
- Build: `npx tsc`
- Run: `npx ts-node src/index.ts`
- Dev (auto-reload): `npx nodemon`

## Dependencies
- [xlsx](https://www.npmjs.com/package/xlsx)
- typescript
- ts-node

## To Do
- Implement CLI argument parsing
- Add merge logic
- Add export functionality

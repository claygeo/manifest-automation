# Auto Manifest

A Node.js script for automating inventory transfer creation from Excel files, using Playwright for browser automation and Electron for file selection and confirmation dialogs. It parses Excel files, validates transfer data, and automates transfer creation on a web-based platform (e.g., SweedPos).

## Features
- Electron-based file selection dialog for choosing Excel files
- Excel parsing with validation for transfer details (store, driver, vehicle, dates, products)
- User confirmation via Electron window
- Playwright automation for Microsoft login and transfer creation
- Robust logging, error handling, and progress tracking
- Unit tests for file selection

## Prerequisites
- Node.js and npm
- Playwright browsers (`npx playwright install`)
- Electron for file dialogs
- Environment variables: `MS_USERNAME`, `MS_PASSWORD` in a `.env` file

## Setup
1. Clone the repository: `git clone [your-repo-url]`
2. Navigate to the project directory: `cd auto-manifest`
3. Install dependencies: `npm install`
4. Install Playwright browsers: `npx playwright install`
5. Create a `.env` file with Microsoft credentials
6. Run the script: `npm run start`
7. Run tests: `npm run test`

## Notes
- Ensure `.env` and `temp/` are not committed (excluded via `.gitignore`).
- The script is designed for a specific inventory system (e.g., Curaleaf), used with permission.
- Temporary files are stored in `temp/` (e.g., `selectedExcel.json`, `transferData.json`).

const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

const [,, input,filename] = process.argv

// Get absolute paths
const powerShellScriptPath = path.resolve(__dirname, 'contoh.ps1');
// const excelFilePath = path.resolve(__dirname, 'sample.xlsx');
const excelFilePath = path.resolve(__dirname, input);
const pdfFilePath = path.resolve(__dirname,  filename ?? 'output.pdf');

// Verify files exist before proceeding
if (!fs.existsSync(powerShellScriptPath)) {
    console.error(`PowerShell script not found at: ${powerShellScriptPath}`);
    process.exit(1);
}

if (!fs.existsSync(excelFilePath)) {
    console.error(`Excel file not found at: ${excelFilePath}`);
    process.exit(1);
}

// Use spawn instead of exec for better argument handling
const powershell = spawn('powershell.exe', [
    '-NoProfile',
    '-NonInteractive',
    '-ExecutionPolicy', 'Bypass',
    '-File', powerShellScriptPath,
    excelFilePath,
    pdfFilePath
]);

powershell.stdout.on('data', (data) => {
    console.log('Output:', data.toString());
});

powershell.stderr.on('data', (data) => {
    console.error('PowerShell Errors:', data.toString());
});

powershell.on('error', (error) => {
    console.error('Failed to start PowerShell:', error);
});

powershell.on('close', (code) => {
    if (code !== 0) {
        console.error(`PowerShell process exited with code ${code}`);
    } else {
        if (fs.existsSync(pdfFilePath)) {
            console.log(`PDF successfully created at: ${pdfFilePath}`);
        } else {
            console.error('PDF was not created. Check if Excel is installed and working properly.');
        }
    }
});
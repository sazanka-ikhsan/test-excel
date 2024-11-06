const { exec } = require('child_process');
const path = require('path');

const powerShellScriptPath = path.join(__dirname, 'contoh.ps1'); // Adjust the path as necessary
const excelFilePath = "C:\\x\\exceltopdf\\sample.xlsx"; // Change this to your Excel file path
const pdfFilePath = "C:\\x\\exceltopdf\\output.pdf"; // Change this to your desired PDF output path

// Construct the command to run the PowerShell script
const command = `powershell -ExecutionPolicy Bypass -File "${powerShellScriptPath}" "${excelFilePath}" "${pdfFilePath}"`;

exec(command, (error, stdout, stderr) => {
    if (error) {
        console.error(`Error: ${error.message}`);
        return;
    }
    if (stderr) {
        console.error(`Stderr: ${stderr}`);
        return;
    }
    console.log(`Stdout: ${stdout}`);
});

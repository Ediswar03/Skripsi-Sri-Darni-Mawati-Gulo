const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');

const dirPath = path.join(__dirname, 'Jurnal Referensi');

function getFiles(dir, files = []) {
    const fileList = fs.readdirSync(dir);
    for (const file of fileList) {
        const name = path.join(dir, file);
        if (fs.statSync(name).isDirectory()) {
            getFiles(name, files);
        } else {
            if (name.endsWith('.pdf')) {
                files.push(name);
            }
        }
    }
    return files;
}

async function extractInfo() {
    const pdfFiles = getFiles(dirPath);
    console.log(`Found ${pdfFiles.length} PDF files.`);

    for (const file of pdfFiles) {
        const relativePath = path.relative(dirPath, file);
        console.log(`\n========================================`);
        console.log(`FILE: ${relativePath}`);
        console.log(`========================================`);

        try {
            const dataBuffer = fs.readFileSync(file);
            // Limit extraction to first 6000 characters to keep it fast and small
            const data = await pdf(dataBuffer, {
                max: 2 // read first 2 pages
            });
            
            console.log("TEXT EXCERPT:");
            console.log(data.text.substring(0, 1500));
        } catch (err) {
            console.error(`Error reading ${relativePath}:`, err.message);
        }
    }
}

extractInfo();

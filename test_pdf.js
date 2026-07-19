const fs = require('fs');
const path = require('path');
const pdf = require('pdf-parse');

async function test() {
    const filePath = path.join(__dirname, 'Jurnal Referensi', 'wisnumahendri,+Guntoro.pdf');
    const dataBuffer = fs.readFileSync(filePath);
    const p = new pdf.PDFParse({ data: dataBuffer });
    const textObj = await p.getText();
    console.log("TEXT EXTRACTED SUCCESSFULLY:");
    console.log(textObj.text.substring(0, 500));
}

test().catch(err => console.error("Error:", err));

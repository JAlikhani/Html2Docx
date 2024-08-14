import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import htmlToDocx from 'html-to-docx';

async function convertHtmlToDocx(htmlContent, outputPath) {
    try {
        const docxContent = await htmlToDocx(htmlContent);
        
        console.log(docxContent)
        fs.writeFileSync(outputPath, docxContent);
        console.log("DOCX file created successfully!");
    } catch (error) {
        console.error("Error creating DOCX file 6:", error);
    }
}

const htmlContent = `
    <h1>Hello World</h1>
    <p>This is a paragraph in the document.</p>
`;

const outputPath = "output.docx";

convertHtmlToDocx(htmlContent, outputPath);
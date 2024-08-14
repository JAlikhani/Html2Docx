import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import htmlToDocx from 'html-to-docx';


async function convertHtmlToDocx(htmlContent, outputPath) {
    const docxContent = await htmlToDocx(htmlContent);
    const zip = new PizZip(docxContent);
    const doc = new Docxtemplater(zip);
    const docBuf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputPath, docBuf);
}

const htmlContent = `
<h1 style="color:blue">Page 12</h1>
`;

const outputPath = "output.docx";

convertHtmlToDocx(htmlContent, outputPath)
    .then(() => console.log("DOCX file created successfully!"))
    .catch((error) => console.error("Error creating DOCX file:", error));

    async function convertHtmlToDocx1(htmlSections, outputPath) {
        try {
            // Add each HTML section with page breaks in between
            let fullDoc = '';
            for (const htmlContent of htmlSections) {
                fullDoc += htmlContent + '<br style="page-break-before: always;">';
            }
    
            const docxContent = await htmlToDocx(fullDoc, {
                pageSize: { width: 12240, height: 15840 }, // A4 page size in twips (1/20 of a point)
                margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1-inch margins
            });
    
            console.log(fullDoc)
            fs.writeFileSync(outputPath, docxContent);
            console.log("DOCX file created successfully!");
        } catch (error) {
            console.error("Error creating DOCX file:", error);
        }
    }
    
    // Example HTML content for multiple pages
    const htmlSections = [
        `   <h1>Page 11</h1>
            <p>This is the content for the first page.</p>
        `,
        `
            <h1>Page 2</h1>
            <p>This is the content for the second page.</p>
        `,
        `
            <h1>Page 3</h1>
            <p>This is the content for the third page.</p>
        `
    ];
    
    const outputPath1 = "output1.docx";
    convertHtmlToDocx1(htmlSections, outputPath1);
import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import htmlToDocx from 'html-to-docx';


async function convertHtmlToDocx(htmlSections, templatePath, outputPath) {
    try {
        // Load the DOCX template
        const templateFile = fs.readFileSync(templatePath, 'binary');
        const zip = new PizZip(templateFile);
        const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

        // Convert each HTML section to DOCX and concatenate
        let fullDoc = '';
        for (const htmlContent of htmlSections) {
            const docxContent = await htmlToDocx(htmlContent, {
                pageSize: { width: 12240, height: 15840 }, // A4 page size in twips (1/20 of a point)
                margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 }, // 1-inch margins
            });
            fullDoc += docxContent.toString('utf8') + '<br style="page-break-before: always;">';
        }

        //console.log(fullDoc)
        // Set the template's content with the converted HTML
        doc.setData({ content: fullDoc });

        // Render the document
        doc.render( {cellBackground: "#00ff00",
    textColor: "white",
    textAlign: "center", // justify,center,left,right
    pStyle: "Heading",
    fontSize: 22,
    fontFamily: "Calibri",first_name: "John",
            last_name: "Doe", first_name_1: "22222John",
            last_name_1: "m,,,,Doe",});

        // Generate the DOCX file and save it
        const buf = doc.getZip().generate({ type: 'nodebuffer' });
        fs.writeFileSync(outputPath, buf);
        console.log("DOCX file created successfully!");
    } catch (error) {
        console.error("Error creating DOCX file:", error);
    }
}

// Example HTML content for multiple pages
const htmlSections = [
    `
        <h1>Page 1</h1>
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

const templatePath = "in8.docx";
const outputPath = "output.docx";
convertHtmlToDocx(htmlSections, templatePath, outputPath);
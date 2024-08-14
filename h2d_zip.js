import fs from 'fs';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import htmlToDocx from 'html-to-docx';


async function convertHtmlToDocx(htmlContent, outputPath) {
    const docxContent = await htmlToDocx(htmlContent);
    const zip = new PizZip(docxContent);
    const doc = new Docxtemplater(zip, { paragraphLoop: true });
    const docBuf = doc.getZip().generate({ type: 'nodebuffer' });
    fs.writeFileSync(outputPath, docBuf);
}

const htmlContent = `
    <h1>Hello World 2</h1>
    <p>This is a paragraph in the document.</p>
   <p4>Author Is: Mehrnoosh Alikhani</p4>
   <button>Kairim</button>
   <p>I am normal</p>
<p style="color:red;">I am red</p>
<p style="color:blue;">I am blue</p>
<p style="font-size:50px;">I am big</p>
<h1 style="background-color:powderblue;">This is a heading</h1>
<p style="background-color:tomato;">This is a paragraph.</p>


`;

const outputPath = "output.docx";

convertHtmlToDocx(htmlContent, outputPath)
    .then(() => console.log("DOCX file created successfully!"))
    .catch((error) => console.error("Error creating DOCX file:", error));
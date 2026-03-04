import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';

export default async function handler(req, res) {
  // Only POST
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  try {
    const { templateBase64, data } = req.body;

    if (!templateBase64 || !data) {
      return res.status(400).json({ 
        error: 'Dati mancanti', 
        received: Object.keys(req.body) 
      });
    }

    // Base64 to Buffer
    const templateBuffer = Buffer.from(templateBase64, 'base64');

    // PizZip and Docxtemplater
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Insert data in the template and render the document
    doc.setData(data);

    doc.render();

    // Generator 
    const outputBuffer = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=CV_Generato.docx');
    
    return res.send(outputBuffer);

  } catch (error) {
    console.error("Errore Generazione:", error);
    return res.status(500).json({ 
      error: 'Errore interno durante la generazione', 
      details: error.message 
    });
  }
}

import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';

export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).send('Method Not Allowed');
  }

  try {
    console.log("Body ricevuto:", Object.keys(req.body)); 
    
    const { templateBase64, data } = req.body;
    const { templateBase64, data } = req.body;

    if (!templateBase64 || !data) {
      return res.status(400).send('Missing templateBase64 or data');
    }

    // Converti il template da Base64 a Buffer
    const templateBuffer = Buffer.from(templateBase64, 'base64');

    // Carica il template
    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

    // Inserisci i dati
    doc.setData(data);

    // Renderizza
    doc.render();

    // Genera il file finale
    const outputBuffer = doc.getZip().generate({
      type: 'nodebuffer',
      compression: 'DEFLATE',
    });

    // Restituisci il file come stream binario
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=CV_Generato.docx');
    return res.send(outputBuffer);

  } catch (error) {
    console.error(error);
    return res.status(500).json({ error: 'Errore durante la generazione del documento', details: error.message });
  }
}

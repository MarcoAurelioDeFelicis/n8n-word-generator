import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import expressions from 'angular-expressions';
import _ from 'lodash';

// Registrazione dei filtri con correzione per lodash
// Utilizziamo String(s).toLowerCase() nativo invece di _.toLowerCase
expressions.filters.lowerCase = (s) => (s ? String(s).toLowerCase() : "");

// ucWords: trasforma "MARCO AURELIO" in "Marco Aurelio"
expressions.filters.ucWords = (s) => (s ? _.startCase(String(s).toLowerCase()) : "");

// arrayJoin: unisce elementi di un array con un separatore
expressions.filters.arrayJoin = (arr, sep) => (Array.isArray(arr) ? arr.join(sep) : arr);

// convCRLF: gestore per i ritorni a capo (necessario per Docxtemplater)
expressions.filters.convCRLF = (s) => s;

function expressionsParser(tag) {
    if (!tag) return { get: (s) => s };
    
    try {
        // Gestione delle virgolette intelligenti che Word spesso inserisce
        const expr = expressions.compile(tag.replace(/(’|“|”|‘)/g, "'"));
        return {
            get: function(scope, context) {
                let obj = null;
                const scopeList = context.scopeList;
                const num = context.num;
                // Risale la gerarchia degli scope per trovare il valore
                for (let i = num; i >= 0; i--) {
                    obj = expr(scopeList[i]);
                    if (obj !== undefined && obj !== null) return obj;
                }
                return obj;
            },
        };
    } catch (e) {
        console.error("Errore compilazione tag:", tag, e);
        return { get: () => "{ERRORE TAG: " + tag + "}" };
    }
}

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

    try {
        const { templateBase64, data } = req.body;
        if (!templateBase64 || !data) throw new Error("Dati mancanti (templateBase64 o data)");

        const payload = typeof data === 'string' ? JSON.parse(data) : data;
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        const zip = new PizZip(templateBuffer);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            parser: expressionsParser
        });

        // Applica i dati al template (avvolti nell'oggetto "d" come nel tuo Word)
        doc.render({ d: payload });

        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=CV_Generato.docx');
        return res.send(outputBuffer);

    } catch (error) {
        console.error("Errore durante la generazione:", error);
        // Restituisce l'errore dettagliato per il debug in n8n
        return res.status(500).json({ 
            error: error.message: error.message // Questo aiuta a vedere quale tag specifico ha fallito
        });
    }
}

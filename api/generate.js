import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import expressions from 'angular-expressions';
import _ from 'lodash';

// Registrazione dei filtri
expressions.filters.lowerCase = (s) => (s ? s.toLowerCase() : "");
expressions.filters.ucWords = (s) => (s ? _.startCase(_.toLowerCase(s)) : "");
expressions.filters.arrayJoin = (arr, sep) => (Array.isArray(arr) ? arr.join(sep) : arr);

// Parser per gestire la sintassi standard Angular |
function expressionsParser(tag) {
    if (!tag) return { get: (s) => s };
    
    try {
        const expr = expressions.compile(tag.replace(/(’|“|”|‘)/g, "'"));
        return {
            get: function(scope, context) {
                let obj = null;
                const scopeList = context.scopeList;
                const num = context.num;
                for (let i = num; i >= 0; i--) {
                    obj = expr(scopeList[i]);
                    if (obj !== undefined && obj !== null) return obj;
                }
                return obj;
            },
        };
    } catch (e) {
        return { get: () => "{ERRORE TAG: " + tag + "}" };
    }
}

export default async function handler(req, res) {
    if (req.method !== 'POST') return res.status(405).send('Method Not Allowed');

    try {
        const { templateBase64, data } = req.body;
        if (!templateBase64 || !data) throw new Error("Dati mancanti");

        const payload = typeof data === 'string' ? JSON.parse(data) : data;
        const templateBuffer = Buffer.from(templateBase64, 'base64');
        const zip = new PizZip(templateBuffer);

        const doc = new Docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            parser: expressionsParser
        });

        // Esegue il rendering (avvolgiamo i dati in 'd' per il template)
        doc.render({ d: payload });

        const outputBuffer = doc.getZip().generate({
            type: 'nodebuffer',
            compression: 'DEFLATE',
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', 'attachment; filename=CV_Generato.docx');
        return res.send(outputBuffer);

    } catch (error) {
        console.error("Errore:", error.message);
        return res.status(500).json({ error: error.message });
    }
}

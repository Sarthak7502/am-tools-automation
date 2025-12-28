const express = require('express');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();

app.use(express.static('public'));

// Increase limit if you send large text data
app.use(express.json({ limit: '50mb' })); 
app.use(express.urlencoded({ extended: true }));

app.post('/generate', (req, res) => {
    try {
        console.log("--- STARTING GENERATION (TEXT ONLY) ---");

        // 1. Get User Input
        // Note: Ensure your frontend sends 'Content-Type: application/json'
        // If your frontend sends stringified JSON inside a key (like 'jsonData'), use: JSON.parse(req.body.jsonData)
        let userInput = req.body;
        
        // Safety check: if data came in as a stringified field named jsonData
        if (req.body.jsonData) {
            userInput = JSON.parse(req.body.jsonData);
        }

        const type = req.body.docType || userInput.docType;

        // 2. Map Data (Text Fields Only)
        const sessionsData = (userInput.sessions || []).map((session) => {
            return {
                session_title: session.session_title || "",
                grades: session.grades || "",
                // Clean formatting of summary if needed
                summary: (session.summary || "").replace(/\r?\n/g, " "), 
                sdgs: (session.sdgs || []).map(sdg => ({
                    number: sdg.number || "",
                    name: sdg.name || ""
                }))
            };
        });

        const data = {
            ...userInput,
            school_name: userInput.school_name || "Document",
            report_period: userInput.report_period || "",
            activation_pct: userInput.activation_pct || "",
            // Inject Globals for logic safety
            req_sessions: userInput.req_sessions || "",
            req_month: userInput.req_month || "",
            sessions: sessionsData
        };

        // 3. Load Template
        const templatePath = type === 'MOU' ? './templates/mou_template.docx' : './templates/report_template.docx';
        const content = fs.readFileSync(path.resolve(__dirname, templatePath), 'binary');
        const zip = new PizZip(content);

        // 4. Configure Engine (No Image Module)
        const doc = new Docxtemplater(zip, {
            delimiters: { start: '%%%', end: '%%%' },
            // Removed modules: [] 
            paragraphLoop: true,
            linebreaks: true,
            nullGetter: () => "" 
        });

        // 5. Render & Generate
        doc.render(data);

        const buf = doc.getZip().generate({ type: 'nodebuffer', compression: 'DEFLATE' });
        const safeName = (data.school_name).replace(/[^a-z0-9]/gi, '_');
        const fileName = type === 'REPORT' ? `Progress_Report_${safeName}.docx` : `MOU_${safeName}.docx`;

        res.set({
            'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'Content-Disposition': `attachment; filename="${fileName}"`
        });
        
        res.send(buf);
        console.log("--- SUCCESS ---");

    } catch (error) {
        console.error("SERVER ERROR:", error);
        res.status(500).send("Error: " + error.message);
    }
});

app.listen(3000, () => console.log('Server running at http://localhost:3000'));
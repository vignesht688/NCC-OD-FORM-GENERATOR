const express = require('express');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const docxLib = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, VerticalAlign, BorderStyle, XmlComponent } = docxLib;

// ============================================================
// HELPER: Build a proper vMerge element
// For OpenXML: RESTART = <w:vMerge w:val="restart"/>
//              CONTINUE = <w:vMerge/>  (NO val attribute!)
// ============================================================
class VMergeContinue extends XmlComponent {
    constructor() {
        super('w:vMerge'); // empty element, no val = CONTINUE
    }
}

// ============================================================
// SHARED: Build OD form table rows
// ============================================================
function buildODTableRows(cadets, particularsVal, odReasonVal) {
    const rows = [];
    // Header row
    rows.push(new TableRow({
        children: [
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'S.NO',        font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'REGISTER NO', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'NAME',        font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'DEPARTMENT',  font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'SHIFT',       font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'PARTICULARS', font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'ON DUTY',     font: 'Cambria', size: 24, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
        ]
    }));

    const N = cadets.length;
    let P = 1;
    if (N > 27) {
        P = Math.ceil((N - 27) / 32) + 1;
    }
    const midPage = Math.ceil(P / 2);
    const startRow = midPage === 1 ? 0 : 27 + (midPage - 2) * 32;

    // Data rows
    cadets.forEach((cadet, index) => {
        let isPartRestart = false;
        let partVal = "";
        let isOdRestart = false;
        let odVal = "";
        
        if (startRow > 0) {
            if (index === 0) {
                isPartRestart = true;
                partVal = "";
                isOdRestart = true;
                odVal = "";
            } else if (index === startRow) {
                isPartRestart = true;
                partVal = particularsVal;
                isOdRestart = true;
                odVal = odReasonVal;
            }
        } else {
            if (index === 0) {
                isPartRestart = true;
                partVal = particularsVal;
                isOdRestart = true;
                odVal = odReasonVal;
            }
        }

        const partMerge = startRow > 0
            ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
            : (index === 0 ? 'restart' : 'continue');

        const odMerge = startRow > 0
            ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
            : (index === 0 ? 'restart' : 'continue');

        const partCell = new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isPartRestart ? partVal : "", font: 'Cambria', size: 24 })] })],
            verticalAlign: VerticalAlign.CENTER,
            verticalMerge: partMerge
        });

        const odCell = new TableCell({
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isOdRestart ? odVal : "", font: 'Cambria', size: 24 })] })],
            verticalAlign: VerticalAlign.CENTER,
            verticalMerge: odMerge
        });

        // Fix continue cells: remove wrong val attribute, keep element empty
        if (partMerge === 'continue') {
            const tcPr = partCell.root.find(el => el && el.rootKey === 'w:tcPr');
            if (tcPr) {
                const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                if (vMergeEl) vMergeEl.root = [];
            }
        }
        if (odMerge === 'continue') {
            const tcPr = odCell.root.find(el => el && el.rootKey === 'w:tcPr');
            if (tcPr) {
                const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                if (vMergeEl) vMergeEl.root = [];
            }
        }

        rows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1),             font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || '',        font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: ' ' + cadet.name,              font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: ' ' + (cadet.department || ''), font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || '',              font: 'Cambria', size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                partCell,
                odCell,
            ]
        }));
    });
    return rows;
}

// Helper to sort cadets by year (III first, then II, then I) and name
function getYearWeight(department) {
    const dept = (department || '').trim().toUpperCase();
    if (dept.startsWith('III')) return 3;
    if (dept.startsWith('II')) return 2;
    if (dept.startsWith('I')) return 1;
    return 0;
}

function sortCadetsForOD(a, b) {
    // 1. Sort by Year (III first, then II, then I)
    const weightA = getYearWeight(a.department);
    const weightB = getYearWeight(b.department);
    if (weightB !== weightA) {
        return weightB - weightA;
    }

    // 2. Sort by Name (alphabetically)
    const nameA = (a.name || '').trim().toUpperCase();
    const nameB = (b.name || '').trim().toUpperCase();
    if (nameA !== nameB) {
        return nameA.localeCompare(nameB);
    }

    // 3. Fallback to Register Number
    const regA = (a.register_no || '').trim();
    const regB = (b.register_no || '').trim();
    return regA.localeCompare(regB);
}

// ============================================================
// SHARED: Build complete OD Document
// ============================================================
async function buildODDocument(cadets, particularsVal, odReasonVal, docConfig) {
    if (Array.isArray(cadets)) {
        cadets.sort(sortCadetsForOD);
    }
    const college  = (docConfig && docConfig.college)   || 'SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)';
    const unit     = (docConfig && docConfig.unit)      || 'NATIONAL CADET CORPS';
    const formType = (docConfig && docConfig.formType)  || 'OD FORM';
    const sigLeft  = (docConfig && docConfig.sigLeft)   || 'NCC Officer';
    const sigRight = (docConfig && docConfig.sigRight)  || 'Principal';

    const tableRows = buildODTableRows(cadets, particularsVal, odReasonVal);

    const doc = new Document({
        sections: [{
            properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
            children: [
                new Paragraph({ alignment: AlignmentType.LEFT,   spacing: { after: 120 }, children: [new TextRun({ text: college,  font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: unit,     font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: formType, font: 'Bookman Old Style', bold: true, size: 24 })] }),
                new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                new Paragraph({ spacing: { before: 2160 }, children: [new TextRun('')] }),
                new Table({
                    width: { size: 100, type: WidthType.PERCENTAGE },
                    borders: {
                        top:             { style: BorderStyle.NONE, size: 0 },
                        bottom:          { style: BorderStyle.NONE, size: 0 },
                        left:            { style: BorderStyle.NONE, size: 0 },
                        right:           { style: BorderStyle.NONE, size: 0 },
                        insideHorizontal:{ style: BorderStyle.NONE, size: 0 },
                        insideVertical:  { style: BorderStyle.NONE, size: 0 },
                    },
                    rows: [new TableRow({ children: [
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,  children: [new TextRun({ text: sigLeft,  font: 'Bookman Old Style', size: 24 })] })] }),
                        new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sigRight, font: 'Bookman Old Style', size: 24 })] })] }),
                    ]})]  
                })
            ]
        }]
    });
    return await Packer.toBuffer(doc);
}

const app = express();
const PORT = process.env.PORT || 3000;
const CADETS_FILE = path.join(__dirname, 'cadets.json');
const ATTENDANCE_FILE = path.join(__dirname, 'attendance.json');

// Middleware to parse JSON and URL-encoded bodies
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// Serve static files from the current directory with caching disabled for local development comfort
app.use(express.static(__dirname, {
    setHeaders: (res, path) => {
        res.set('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    }
}));

// Endpoint to fetch cadets from cadets.json
app.get('/api/cadets', (req, res) => {
    fs.readFile(CADETS_FILE, 'utf8', (err, data) => {
        if (err) {
            if (err.code === 'ENOENT') {
                // If file doesn't exist, return empty database
                return res.json([]);
            }
            console.error('Error reading database file:', err);
            return res.status(500).json({ error: 'Failed to read database file' });
        }
        try {
            const json = JSON.parse(data || '[]');
            res.json(json);
        } catch (parseErr) {
            console.error('Error parsing database JSON:', parseErr);
            res.status(500).json({ error: 'Invalid database format' });
        }
    });
});

// Endpoint to write cadets to cadets.json
app.post('/api/cadets', (req, res) => {
    const cadets = req.body;
    if (!Array.isArray(cadets)) {
        return res.status(400).json({ error: 'Data must be an array of cadets' });
    }

    fs.writeFile(CADETS_FILE, JSON.stringify(cadets, null, 2), 'utf8', (err) => {
        if (err) {
            console.error('Error writing database file:', err);
            return res.status(500).json({ error: 'Failed to save database file' });
        }
        console.log(`Database updated successfully with ${cadets.length} cadets.`);
        res.json({ success: true, count: cadets.length });
    });
});

// Endpoint to generate and download Word document on the server (bypasses browser security limits)
app.post('/api/generate-docx', async (req, res) => {
    try {
        let payload = req.body;
        if (req.body.payload) {
            payload = JSON.parse(req.body.payload);
        }

        const { selectedCadets, particularsVal, odReasonVal, selectedDate, docConfig } = payload;

        if (!selectedCadets || selectedCadets.length === 0) {
            return res.status(400).send("No cadets selected to export.");
        }

        selectedCadets.sort(sortCadetsForOD);

        // Build Table Rows
        const tableRows = [];

        // 1. Table Header Row
        tableRows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ]
        }));

        // 2. Add Cadet Rows
        selectedCadets.forEach((cadet, index) => {
            const rowCells = [
                // Col 0: S.No
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 1: Reg No
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 2: Name (left aligned)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 3: Dept (left aligned)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department ? " " + cadet.department : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 4: Shift
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                // Col 5: Particulars (vertically merged)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? particularsVal : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                    verticalMerge: index === 0 ? "restart" : "continue"
                }),
                // Col 6: On Duty (vertically merged)
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? odReasonVal : "", font: "Cambria", size: 24 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                    verticalMerge: index === 0 ? "restart" : "continue"
                })
            ];

            tableRows.push(new TableRow({ children: rowCells }));
        });

        // Date String helper
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(selectedDate);
        const dayStr = String(d.getDate()).padStart(2, '0');
        const monthStr = months[d.getMonth()];
        const yearStr = d.getFullYear();
        const formattedDateStr = `${dayStr} ${monthStr} ${yearStr}`;

        // 3. Build Document Structure
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                        }
                    }
                },
                children: [
                    // College Header
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.college,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // NCC Subtitle
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.unit,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Title Header
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 360 },
                        children: [
                            new TextRun({
                                text: docConfig.formType,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Main Grid Table
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        rows: tableRows
                    }),
                    // Space before signatures
                    new Paragraph({
                        spacing: { before: 1440 },
                        children: [new TextRun("")]
                    }),
                    // Signatures Grid
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            left: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            right: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.LEFT,
                                            children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating docx on server:", err);
        res.status(500).send("Error generating document: " + err.message);
    }
});

// ============================================================
// /api/generate-and-open  — saves OD form to Desktop & opens it in MS Word
// ============================================================
app.post('/api/generate-and-open', async (req, res) => {
    try {
        if (process.env.VERCEL || process.env.RENDER || process.env.NODE_ENV === 'production') {
            return res.status(400).json({ 
                error: 'The "Open in Word" feature is only supported when running the app locally on your computer. Please use the "Download (.docx)" button instead.' 
            });
        }
        const { selectedCadets, particularsVal, odReasonVal, selectedDate, docConfig: cfg } = req.body;

        if (!selectedCadets || selectedCadets.length === 0) {
            return res.status(400).json({ error: 'No cadets selected.' });
        }

        const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
        const d = new Date(selectedDate);
        const formattedDate = `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;

        const buffer = await buildODDocument(selectedCadets, particularsVal, odReasonVal, cfg);

        // Save to Desktop
        const desktopPath = path.join(require('os').homedir(), 'OneDrive', 'Desktop');
        const filename = `od-form-${formattedDate.toLowerCase().replace(/ /g,'-')}.docx`;
        const filePath = path.join(desktopPath, filename);
        fs.writeFileSync(filePath, buffer);

        // Auto-open in MS Word
        exec(`start "" "${filePath}"`, { shell: true }, (err) => {
            if (err) console.error('Auto-open failed:', err);
        });

        console.log(`OD Form saved and opened: ${filePath}`);
        res.json({ success: true, filename, path: filePath });

    } catch (err) {
        console.error('Error in generate-and-open:', err);
        res.status(500).json({ error: err.message });
    }
});


app.get('/api/attendance', (req, res) => {
    const { date } = req.query;
    if (!date) {
        return res.status(400).json({ error: 'Date is required' });
    }
    fs.readFile(ATTENDANCE_FILE, 'utf8', (err, data) => {
        if (err) {
            if (err.code === 'ENOENT') {
                return res.json({});
            }
            console.error('Error reading attendance database:', err);
            return res.status(500).json({ error: 'Failed to read database' });
        }
        try {
            const db = JSON.parse(data || '{}');
            res.json(db[date] || {});
        } catch (parseErr) {
            console.error('Error parsing database JSON:', parseErr);
            res.status(500).json({ error: 'Invalid database format' });
        }
    });
});

// Endpoint to save attendance for a date
app.post('/api/attendance', (req, res) => {
    const { date, event_name, od_reason, records } = req.body;
    if (!date) {
        return res.status(400).json({ error: 'Date is required' });
    }
    fs.readFile(ATTENDANCE_FILE, 'utf8', (err, data) => {
        let db = {};
        if (!err) {
            try {
                db = JSON.parse(data || '{}');
            } catch (parseErr) {
                console.error('Error parsing database JSON:', parseErr);
            }
        }
        
        db[date] = { event_name, od_reason, records };
        
        fs.writeFile(ATTENDANCE_FILE, JSON.stringify(db, null, 2), 'utf8', (writeErr) => {
            if (writeErr) {
                console.error('Error writing attendance database:', writeErr);
                return res.status(500).json({ error: 'Failed to save database' });
            }
            res.json({ success: true });
        });
    });
});

// Endpoint to fetch attendance history summary
app.get('/api/attendance/history', (req, res) => {
    fs.readFile(ATTENDANCE_FILE, 'utf8', (err, data) => {
        if (err) {
            if (err.code === 'ENOENT') {
                return res.json([]);
            }
            console.error('Error reading attendance database:', err);
            return res.status(500).json({ error: 'Failed to read database' });
        }
        try {
            const db = JSON.parse(data || '{}');
            const summary = [];
            for (const date in db) {
                const record = db[date];
                const counts = { present: 0, absent: 0, od: 0, total: 0 };
                if (record.records) {
                    for (const cadetId in record.records) {
                        const status = record.records[cadetId];
                        counts.total++;
                        if (status === 'present') counts.present++;
                        else if (status === 'absent') counts.absent++;
                        else if (status === 'od') counts.od++;
                    }
                }
                summary.push({
                    date,
                    event_name: record.event_name || 'Regular parade',
                    od_reason: record.od_reason || '',
                    counts
                });
            }
            // Sort by date descending
            summary.sort((a, b) => b.date.localeCompare(a.date));
            res.json(summary);
        } catch (parseErr) {
            console.error('Error parsing database JSON:', parseErr);
            res.status(500).json({ error: 'Invalid database format' });
        }
    });
});

// Endpoint to delete attendance for a date
app.delete('/api/attendance', (req, res) => {
    const { date } = req.query;
    if (!date) {
        return res.status(400).json({ error: 'Date is required' });
    }
    fs.readFile(ATTENDANCE_FILE, 'utf8', (err, data) => {
        if (err) {
            console.error('Error reading database file:', err);
            return res.status(500).json({ error: 'Failed to read database' });
        }
        try {
            const db = JSON.parse(data || '{}');
            if (db[date]) {
                delete db[date];
                fs.writeFile(ATTENDANCE_FILE, JSON.stringify(db, null, 2), 'utf8', (writeErr) => {
                    if (writeErr) {
                        return res.status(500).json({ error: 'Failed to save database' });
                    }
                    res.json({ success: true });
                });
            } else {
                res.json({ success: true, message: 'No record found to delete' });
            }
        } catch (parseErr) {
            res.status(500).json({ error: 'Invalid database format' });
        }
    });
});

// Endpoint to generate attendance sheet docx
app.post('/api/generate-attendance-sheet-docx', async (req, res) => {
    try {
        let payload = req.body;
        if (req.body.payload) {
            payload = JSON.parse(req.body.payload);
        }

        const { date, records, docConfig } = payload;

        if (!records || records.length === 0) {
            return res.status(400).send("No records to export.");
        }

        // Build Table Rows
        const tableRows = [];

        // 1. Table Header Row
        tableRows.push(new TableRow({
            children: [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ]
        }));

        // 2. Add Cadet Records Rows
        records.forEach((cadet, index) => {
            const statusText = (cadet.status || "PRESENT").toUpperCase();
            const rowCells = [
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department ? " " + cadet.department : "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 22 })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
                new TableCell({
                    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusText, font: "Cambria", size: 22, bold: true })] })],
                    verticalAlign: VerticalAlign.CENTER,
                }),
            ];

            tableRows.push(new TableRow({ children: rowCells }));
        });

        // Date String helper
        const months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
        const d = new Date(date);
        const dayStr = String(d.getDate()).padStart(2, '0');
        const monthStr = months[d.getMonth()];
        const yearStr = d.getFullYear();
        const formattedDateStr = `${dayStr} ${monthStr} ${yearStr}`;

        // 3. Build Document Structure
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440,
                            bottom: 1440,
                            left: 1440,
                            right: 1440,
                        }
                    }
                },
                children: [
                    // College Header
                    new Paragraph({
                        alignment: AlignmentType.LEFT,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.college,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // NCC Subtitle
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 120 },
                        children: [
                            new TextRun({
                                text: docConfig.unit,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Title Header
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { after: 360 },
                        children: [
                            new TextRun({
                                text: `ATTENDANCE SHEET - ${formattedDateStr}`,
                                font: "Bookman Old Style",
                                bold: true,
                                size: 24,
                            })
                        ]
                    }),
                    // Main Grid Table
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        rows: tableRows
                    }),
                    // Space before signatures
                    new Paragraph({
                        spacing: { before: 1440 },
                        children: [new TextRun("")]
                    }),
                    // Signatures Grid
                    new Table({
                        width: {
                            size: 100,
                            type: WidthType.PERCENTAGE,
                        },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            left: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            right: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
                            insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
                        },
                        rows: [
                            new TableRow({
                                children: [
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.LEFT,
                                            children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    }),
                                    new TableCell({
                                        children: [new Paragraph({
                                            alignment: AlignmentType.RIGHT,
                                            children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })]
                                        })],
                                    })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `attendance-${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating attendance docx:", err);
        res.status(500).send("Error generating document: " + err.message);
    }
});

// ============================================================
// ATTENDANCE ROUTES
// ============================================================

// Helper: read attendance database
function readAttendanceDB() {
    try {
        if (fs.existsSync(ATTENDANCE_FILE)) {
            return JSON.parse(fs.readFileSync(ATTENDANCE_FILE, 'utf8') || '{}');
        }
    } catch (e) {
        console.error("Error reading attendance DB:", e);
    }
    return {};
}

// Helper: write attendance database
function writeAttendanceDB(data) {
    fs.writeFileSync(ATTENDANCE_FILE, JSON.stringify(data, null, 2), 'utf8');
}

// GET /api/attendance?date=YYYY-MM-DD  - get one day's record
app.get('/api/attendance', (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date query param required' });
    const db = readAttendanceDB();
    const record = db[date] || { date, event_name: '', od_reason: '', records: {} };
    res.json(record);
});

// POST /api/attendance - save or update one day's record
app.post('/api/attendance', (req, res) => {
    const { date, event_name, od_reason, records } = req.body;
    if (!date) return res.status(400).json({ error: 'date required' });
    const db = readAttendanceDB();
    db[date] = { date, event_name: event_name || '', od_reason: od_reason || '', records: records || {} };
    writeAttendanceDB(db);
    console.log(`Attendance saved for ${date}`);
    res.json({ success: true });
});

// GET /api/attendance/history - get summary list of all recorded sessions
app.get('/api/attendance/history', (req, res) => {
    const db = readAttendanceDB();
    const history = Object.keys(db)
        .sort((a, b) => b.localeCompare(a)) // newest first
        .map(date => {
            const rec = db[date];
            const counts = { present: 0, absent: 0, od: 0, total: 0 };
            if (rec.records) {
                for (const key in rec.records) {
                    counts.total++;
                    const s = rec.records[key];
                    if (s === 'present') counts.present++;
                    else if (s === 'absent') counts.absent++;
                    else if (s === 'od') counts.od++;
                }
            }
            return {
                date,
                event_name: rec.event_name,
                od_reason: rec.od_reason,
                records: rec.records,
                counts
            };
        });
    res.json(history);
});

// DELETE /api/attendance?date=YYYY-MM-DD - remove a day's record
app.delete('/api/attendance', (req, res) => {
    const date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date query param required' });
    const db = readAttendanceDB();
    if (db[date]) {
        delete db[date];
        writeAttendanceDB(db);
        console.log(`Attendance record deleted for ${date}`);
        res.json({ success: true });
    } else {
        res.status(404).json({ error: 'No record found for that date' });
    }
});

// ============================================================
// DIRECT DOWNLOAD SHORTCUTS (accessible via browser URL bar)
// ============================================================

// GET /download/attendance?date=YYYY-MM-DD
// Directly generates and downloads the attendance sheet Word doc
app.get('/download/attendance', async (req, res) => {
    try {
        const date = req.query.date || new Date().toISOString().split('T')[0];

        // Load cadets
        let cadets = [];
        try { cadets = JSON.parse(fs.readFileSync(CADETS_FILE, 'utf8') || '[]'); } catch (e) {}

        if (cadets.length === 0) {
            return res.status(400).send('No cadets in database. Please add cadets first.');
        }

        // Load attendance records for this date
        const db = readAttendanceDB();
        const attRecord = db[date] || { records: {} };

        // Build records array with status
        const records = cadets.map(c => ({
            register_no: c.register_no,
            name: c.name,
            department: c.department,
            shift: c.shift,
            status: attRecord.records[c.register_no ? c.register_no.trim() : c.name.trim().toLowerCase()] || 'present'
        }));

        // Date formatting
        const months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
        const d = new Date(date);
        const formattedDateStr = `${String(d.getDate()).padStart(2,'0')} ${months[d.getMonth()]} ${d.getFullYear()}`;

        // Build table rows
        const tableRows = [];

        // Header row
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "STATUS", font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
            ]
        }));

        // Data rows
        records.forEach((cadet, index) => {
            const statusText = (cadet.status || 'PRESENT').toUpperCase();
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: " " + (cadet.department || ""), font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 22 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusText, font: "Cambria", size: 22, bold: true })] })], verticalAlign: VerticalAlign.CENTER }),
                ]
            }));
        });

        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: "SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)", font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: "NATIONAL CADET CORPS", font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: `ATTENDANCE SHEET - ${formattedDateStr}`, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 },
                            left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 },
                        },
                        rows: [new TableRow({ children: [
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,  children: [new TextRun({ text: "NCC Officer", font: "Bookman Old Style", size: 24 })] })] }),
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Principal", font: "Bookman Old Style", size: 24 })] })] }),
                        ]})]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        const filename = `attendance-${formattedDateStr.toLowerCase().replace(/ /g, '-')}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
        res.send(buffer);

    } catch (err) {
        console.error("Error generating direct download attendance:", err);
        res.status(500).send("Error: " + err.message);
    }
});

// ============================================================
// /api/generate-for-date  — Quick-generate OD form for ALL cadets for a date
// ============================================================
app.post('/api/generate-for-date', async (req, res) => {
    try {
        if (process.env.VERCEL || process.env.RENDER || process.env.NODE_ENV === 'production') {
            return res.status(400).json({ 
                error: 'The "Quick Gen" (save to Desktop) feature is only supported when running the app locally on your computer. Please select cadets and use the "Download (.docx)" button instead.' 
            });
        }
        const { date, particulars, odReason } = req.body;
        const selectedDate = date || new Date().toISOString().split('T')[0];

        // Load all cadets
        let allCadets = [];
        try { allCadets = JSON.parse(fs.readFileSync(CADETS_FILE, 'utf8') || '[]'); } catch (e) {}

        if (allCadets.length === 0) {
            return res.status(400).json({ error: 'No cadets in database. Please add cadets first.' });
        }

        // Date formatting
        const months = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
        const monthsShort = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"];
        const d = new Date(selectedDate);
        const day = d.getDate();
        const suffix = day === 1 ? 'ST' : day === 2 ? 'ND' : day === 3 ? 'RD' : 'TH';
        const dateStr = `${day}${suffix} ${months[d.getMonth()]} ${d.getFullYear()}`;
        const formattedDate = `${String(day).padStart(2,'0')} ${monthsShort[d.getMonth()]} ${d.getFullYear()}`;

        const particularsVal = particulars || 'NCC ACTIVITY';
        const odReasonVal = odReason || `EXEMPTION FROM ALL ACADEMIC WORKS AND REGULAR CLASSES ON ${dateStr}`;

        const college  = 'SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)';
        const unit     = 'NATIONAL CADET CORPS';
        const formType = 'OD FORM';
        const sigLeft  = 'NCC Officer';
        const sigRight = 'Principal';

        // Build table rows
        const tableRows = [];

        // Header
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO",        font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME",        font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT",  font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT",       font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY",     font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
            ]
        }));

        const N = allCadets.length;
        let P = 1;
        if (N > 27) {
            P = Math.ceil((N - 27) / 32) + 1;
        }
        const midPage = Math.ceil(P / 2);
        const startRow = midPage === 1 ? 0 : 27 + (midPage - 2) * 32;

        // Data rows
        allCadets.forEach((cadet, index) => {
            let isPartRestart = false;
            let partVal = "";
            let isOdRestart = false;
            let odVal = "";
            
            if (startRow > 0) {
                if (index === 0) {
                    isPartRestart = true;
                    partVal = "";
                    isOdRestart = true;
                    odVal = "";
                } else if (index === startRow) {
                    isPartRestart = true;
                    partVal = particularsVal;
                    isOdRestart = true;
                    odVal = odReasonVal;
                }
            } else {
                if (index === 0) {
                    isPartRestart = true;
                    partVal = particularsVal;
                    isOdRestart = true;
                    odVal = odReasonVal;
                }
            }

            const partMerge = startRow > 0
                ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                : (index === 0 ? 'restart' : 'continue');

            const odMerge = startRow > 0
                ? (index < startRow ? (index === 0 ? 'restart' : 'continue') : (index === startRow ? 'restart' : 'continue'))
                : (index === 0 ? 'restart' : 'continue');

            const partCell = new TableCell({
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isPartRestart ? partVal : "", font: "Cambria", size: 24 })] })],
                verticalAlign: VerticalAlign.CENTER,
                verticalMerge: partMerge
            });

            const odCell = new TableCell({
                children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: isOdRestart ? odVal : "", font: "Cambria", size: 24 })] })],
                verticalAlign: VerticalAlign.CENTER,
                verticalMerge: odMerge
            });

            // Fix continue cells: remove wrong val attribute, keep element empty
            if (partMerge === 'continue') {
                const tcPr = partCell.root.find(el => el && el.rootKey === 'w:tcPr');
                if (tcPr) {
                    const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                    if (vMergeEl) vMergeEl.root = [];
                }
            }
            if (odMerge === 'continue') {
                const tcPr = odCell.root.find(el => el && el.rootKey === 'w:tcPr');
                if (tcPr) {
                    const vMergeEl = tcPr.root.find(el => el && el.rootKey === 'w:vMerge');
                    if (vMergeEl) vMergeEl.root = [];
                }
            }

            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1),              font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "",         font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: " " + cadet.name,               font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,   children: [new TextRun({ text: " " + (cadet.department || ""), font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "",               font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    partCell,
                    odCell,
                ]
            }));
        });

        // Build document
        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT,   spacing: { after: 120 }, children: [new TextRun({ text: college,  font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: unit,     font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: formType, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        borders: {
                            top: { style: BorderStyle.NONE, size: 0 }, bottom: { style: BorderStyle.NONE, size: 0 },
                            left: { style: BorderStyle.NONE, size: 0 }, right: { style: BorderStyle.NONE, size: 0 },
                            insideHorizontal: { style: BorderStyle.NONE, size: 0 }, insideVertical: { style: BorderStyle.NONE, size: 0 },
                        },
                        rows: [new TableRow({ children: [
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT,  children: [new TextRun({ text: sigLeft,  font: "Bookman Old Style", size: 24 })] })] }),
                            new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: sigRight, font: "Bookman Old Style", size: 24 })] })] }),
                        ]})]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);

        // Save to Desktop
        const desktopPath = path.join(require('os').homedir(), 'OneDrive', 'Desktop');
        const filename = `od-form-${formattedDate.toLowerCase().replace(/ /g,'-')}.docx`;
        const filePath = path.join(desktopPath, filename);
        fs.writeFileSync(filePath, buffer);

        // Auto-open in MS Word
        exec(`start "" "${filePath}"`, { shell: true }, (err) => {
            if (err) console.error('Auto-open failed:', err);
        });

        console.log(`Quick OD Form generated for ${allCadets.length} cadets: ${filePath}`);
        res.json({ success: true, filename, path: filePath, cadetCount: allCadets.length });

    } catch (err) {
        console.error("Error in generate-for-date:", err);
        res.status(500).json({ error: err.message });
    }
});

// Serve index.html for the root route
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Start the server (only if not running on serverless Vercel environment)
if (!process.env.VERCEL) {
    app.listen(PORT, () => {
        const url = `http://localhost:${PORT}`;
        console.log(`==================================================`);
        console.log(` NCC OD Form Automator server running at: ${url}`);
        console.log(`==================================================`);

        // Auto-open browser on startup (only locally)
        if (process.platform === 'win32' || process.platform === 'darwin') {
            const startCmd = process.platform === 'darwin' ? 'open' : 'start';
            exec(`${startCmd} ${url}`, (err) => {
                if (err) {
                    console.error('Failed to open browser automatically:', err);
                }
            });
        }
    });
}

module.exports = app;

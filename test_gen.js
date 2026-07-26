const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, WidthType, VerticalAlign, BorderStyle } = require('docx');

async function testGenerate() {
    try {
        const selectedCadets = [
            { register_no: "24104070", name: "AKASH S", department: "III BCOM CA", shift: "SHIFT-II" },
            { register_no: "24115002", name: "ANBUSELVAM P", department: "III BSC MATHS", shift: "SHIFT-I" }
        ];
        const particularsVal = "NCC SELECTION PARADE";
        const odReasonVal = "EXEMPTION FROM CLASSES";
        const selectedDate = "2026-07-08";
        const docConfig = {
            college: "SRI RAMAKRISHNA COLLEGE OF ARTS & SCIENCE (AUTONOMOUS)",
            unit: "NATIONAL CADET CORPS",
            formType: "OD FORM",
            sigLeft: "NCC Officer",
            sigRight: "Principal"
        };

        const tableRows = [];

        // Header
        tableRows.push(new TableRow({
            children: [
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "S.NO", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REGISTER NO", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NAME", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DEPARTMENT", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "SHIFT", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PARTICULARS", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ON DUTY", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
            ]
        }));

        // Rows
        selectedCadets.forEach((cadet, index) => {
            tableRows.push(new TableRow({
                children: [
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(index + 1), font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.register_no || "", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: " " + cadet.name, font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: cadet.department ? " " + cadet.department : "", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: cadet.shift || "", font: "Cambria", size: 24 })] })], verticalAlign: VerticalAlign.CENTER }),
                    new TableCell({
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? particularsVal : "", font: "Cambria", size: 24 })] })],
                        verticalAlign: VerticalAlign.CENTER,
                        verticalMerge: index === 0 ? "restart" : "continue"
                    }),
                    new TableCell({
                        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: index === 0 ? odReasonVal : "", font: "Cambria", size: 24 })] })],
                        verticalAlign: VerticalAlign.CENTER,
                        verticalMerge: index === 0 ? "restart" : "continue"
                    })
                ]
            }));
        });

        // Document
        const doc = new Document({
            sections: [{
                properties: { page: { margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
                children: [
                    new Paragraph({ alignment: AlignmentType.LEFT, spacing: { after: 120 }, children: [new TextRun({ text: docConfig.college, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: docConfig.unit, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 360 }, children: [new TextRun({ text: docConfig.formType, font: "Bookman Old Style", bold: true, size: 24 })] }),
                    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }),
                    new Paragraph({ spacing: { before: 1440 }, children: [new TextRun("")] }),
                    new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
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
                                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [new TextRun({ text: docConfig.sigLeft, font: "Bookman Old Style", size: 24 })] })] }),
                                    new TableCell({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: docConfig.sigRight, font: "Bookman Old Style", size: 24 })] })] })
                                ]
                            })
                        ]
                    })
                ]
            }]
        });

        const buffer = await Packer.toBuffer(doc);
        fs.writeFileSync('test_output.docx', buffer);
        console.log('SUCCESS: test_output.docx generated.');
    } catch (err) {
        console.error('ERROR during generation:', err);
    }
}

testGenerate();

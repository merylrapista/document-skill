const docx = require("docx");
const fs = require("fs");

const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, AlignmentType } = docx;

const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({
                text: "Research Report: The Mystery Genre",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "An exploration of core elements, history, and the impact of mystery fiction.",
                        italics: true,
                        size: 28,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
            }),
            new Paragraph({
                text: "1. Introduction",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 },
            }),
            new Paragraph({
                text: "Mystery fiction is a loosely defined genre of writing that involves a crime or a puzzle being solved. The core of a mystery story is a secret that needs to be uncovered, usually revolving around a crime such as murder, theft, or kidnapping.",
                spacing: { after: 200 },
            }),
            new Paragraph({
                text: "2. Core Elements of a Mystery Story",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Paragraph({
                text: "A successful mystery typically relies on a well-established set of elements that keep the reader engaged:",
                spacing: { after: 100 },
            }),
            new Paragraph({ text: "• The Hook: A compelling crime or puzzle that starts the narrative.", bullet: { level: 0 } }),
            new Paragraph({ text: "• The Sleuth: A detective, amateur or professional, who leads the investigation.", bullet: { level: 0 } }),
            new Paragraph({ text: "• The Suspects: A cast of characters with motives and means.", bullet: { level: 0 } }),
            new Paragraph({ text: "• Clues and Red Herrings: True hints mixed with false trails to mislead the sleuth.", bullet: { level: 0 } }),
            new Paragraph({ text: "• The Reveal: The climax where the truth is finally uncovered.", bullet: { level: 0 }, spacing: { after: 200 } }),
            new Paragraph({
                text: "3. Historical Milestones",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: "Era / Year", bold: true })] }),
                            new TableCell({ children: [new Paragraph({ text: "Author", bold: true })] }),
                            new TableCell({ children: [new Paragraph({ text: "Significant Work", bold: true })] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1841")] }),
                            new TableCell({ children: [new Paragraph("Edgar Allan Poe")] }),
                            new TableCell({ children: [new Paragraph("'The Murders in the Rue Morgue'")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1887")] }),
                            new TableCell({ children: [new Paragraph("Arthur Conan Doyle")] }),
                            new TableCell({ children: [new Paragraph("'A Study in Scarlet'")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1920s - 1930s")] }),
                            new TableCell({ children: [new Paragraph("Agatha Christie")] }),
                            new TableCell({ children: [new Paragraph("The 'Golden Age' of Detective Fiction")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1930s - 1940s")] }),
                            new TableCell({ children: [new Paragraph("Raymond Chandler & Dashiell Hammett")] }),
                            new TableCell({ children: [new Paragraph("Hardboiled and Noir Fiction")] }),
                        ],
                    }),
                ],
            }),
            new Paragraph({
                text: "4. Common Sub-genres",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 },
            }),
            new Paragraph({ text: "• Cozy Mysteries: Mild violence, amateur sleuths, often set in small towns.", bullet: { level: 0 } }),
            new Paragraph({ text: "• Hardboiled: Gritty realism, flawed detectives, often set in urban environments.", bullet: { level: 0 } }),
            new Paragraph({ text: "• Police Procedurals: Focuses on the realistic methods of law enforcement.", bullet: { level: 0 } }),
            new Paragraph({ text: "• Locked Room Mysteries: A crime committed under seemingly impossible circumstances.", bullet: { level: 0 }, spacing: { after: 200 } }),
            new Paragraph({
                text: "5. Conclusion",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Paragraph({
                text: "The mystery genre continues to thrive because it taps into the fundamental human desire to solve puzzles and see justice served. Through intricate plotting and memorable characters, mystery writers craft compelling narratives that stand the test of time.",
            }),
        ],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("Mystery_Story_Research.docx", buffer);
    console.log("Document successfully created: Mystery_Story_Research.docx");
});

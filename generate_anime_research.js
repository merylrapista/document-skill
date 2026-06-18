const docx = require("docx");
const fs = require("fs");

const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, BorderStyle, WidthType, AlignmentType } = docx;

const doc = new Document({
    sections: [{
        properties: {},
        children: [
            new Paragraph({
                text: "Research Report: Anime",
                heading: HeadingLevel.HEADING_1,
                alignment: AlignmentType.CENTER,
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "An in-depth look at Japanese animation, its history, and global impact.",
                        italics: true,
                        size: 28,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 },
            }),
            new Paragraph({
                text: "1. Introduction to Anime",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 },
            }),
            new Paragraph({
                text: "Anime refers specifically to animation produced in Japan. Outside of Japan, the term has become a catch-all for the distinct visual and narrative style associated with Japanese animation. It is characterized by vibrant characters, colorful artwork, and fantastical themes.",
                spacing: { after: 200 },
            }),
            new Paragraph({
                text: "2. Key Characteristics",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Paragraph({
                text: "While anime styles vary wildly, some common elements include:",
                spacing: { after: 100 },
            }),
            new Paragraph({
                text: "• Exaggerated physical features, such as large eyes and colorful hair.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "• High emotion and dramatic pacing.",
                bullet: { level: 0 }
            }),
            new Paragraph({
                text: "• Complex character development and intricate storylines, often spanning across many episodes.",
                bullet: { level: 0 },
                spacing: { after: 200 },
            }),
            new Paragraph({
                text: "3. Historical Timeline",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Table({
                width: { size: 100, type: WidthType.PERCENTAGE },
                rows: [
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph({ text: "Era", bold: true })] }),
                            new TableCell({ children: [new Paragraph({ text: "Significant Milestones", bold: true })] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1910s - 1940s")] }),
                            new TableCell({ children: [new Paragraph("Early experiments with animation in Japan; propaganda shorts during WWII.")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1960s")] }),
                            new TableCell({ children: [new Paragraph("Osamu Tezuka's 'Astro Boy' established the modern anime look and broadcast format.")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("1980s - 1990s")] }),
                            new TableCell({ children: [new Paragraph("The Golden Age; rise of Studio Ghibli, 'Akira', 'Neon Genesis Evangelion', and global exports like 'Dragon Ball' and 'Sailor Moon'.")] }),
                        ],
                    }),
                    new TableRow({
                        children: [
                            new TableCell({ children: [new Paragraph("2000s - Present")] }),
                            new TableCell({ children: [new Paragraph("Streaming era, widespread international popularity, and highly diverse genres.")] }),
                        ],
                    }),
                ],
            }),
            new Paragraph({
                text: "4. Major Genres",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 400, after: 200 },
            }),
            new Paragraph({
                text: "Anime spans many genres catering to all demographics. Some notable genres include:",
                spacing: { after: 100 },
            }),
            new Paragraph({ text: "• Shonen: Aimed at young males (e.g., Naruto, One Piece).", bullet: { level: 0 } }),
            new Paragraph({ text: "• Shojo: Aimed at young females (e.g., Sailor Moon, Fruits Basket).", bullet: { level: 0 } }),
            new Paragraph({ text: "• Seinen & Josei: Targeted at adult audiences with more mature themes.", bullet: { level: 0 } }),
            new Paragraph({ text: "• Mecha: Focused on giant robots and sci-fi themes (e.g., Gundam).", bullet: { level: 0 } }),
            new Paragraph({ text: "• Isekai: Protagonists transported to another world.", bullet: { level: 0 }, spacing: { after: 200 } }),
            new Paragraph({
                text: "5. Conclusion",
                heading: HeadingLevel.HEADING_2,
                spacing: { before: 200, after: 200 },
            }),
            new Paragraph({
                text: "Anime has evolved from a niche cultural export into a dominant force in global entertainment. Its unique storytelling techniques and artistic styles continue to influence media, fashion, and art worldwide.",
            }),
        ],
    }],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("Anime_Research.docx", buffer);
    console.log("Document successfully created: Anime_Research.docx");
});

# Skill: docx — .docx Document Generator

Generate Microsoft Word (.docx) documents using JavaScript/Node.js with the `docx` library (by dolanmiu).

## Requirements

Install Node.js (v14+) and the dependency:

```bash
npm install docx
```

## Usage

When the user asks to create a document:

1. Parse their request to determine content (headings, paragraphs, tables, lists, images, styling).
2. Write a JavaScript script that uses the `docx` library to build the document.
3. Run the script with Node.js to produce the `.docx` file.
4. Return the file path to the user.

## Script template

```javascript
const docx = require("docx");
const fs = require("fs");

const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = docx;

const doc = new Document({
  sections: [{
    properties: {},
    children: [
      new Paragraph({
        text: "Title",
        heading: HeadingLevel.TITLE,
      }),
      new Paragraph({
        children: [new TextRun("Body text here.")],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("output.docx", buffer);
  console.log("Document saved as output.docx");
});
```

## Common imports

```javascript
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, Table, TableRow, TableCell, WidthType,
  BorderStyle, PageOrientation, convertInchesToTwip,
} = require("docx");
```

## Styling helpers

| Style | Method |
|-------|--------|
| Bold / Italic | `new TextRun({ text: "...", bold: true, italics: true })` |
| Font size | `new TextRun({ text: "...", size: 24 })` (half-points: 24 = 12pt) |
| Font color | `new TextRun({ text: "...", color: "FF0000" })` |
| Alignment | `new Paragraph({ ..., alignment: AlignmentType.CENTER })` |
| Headings | `heading: HeadingLevel.HEADING_1` (TITLE, HEADING_1-6) |
| Page margins | `page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }` (twips) |
| Tables | `new Table({ rows: [new TableRow({ children: [new TableCell({ children: [new Paragraph("...")] })] })] })` |

## Saving the file

```javascript
Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("output.docx", buffer);
});
```

## File output

Save documents into the working directory with a descriptive filename.

const { Packer } = require("docx");
const fs = require("fs");
const path = require("path");
const { createFixedTable, chunkArray } = require("./docxHelpers"); // You'll split docx logic here
const { Document, Paragraph, TextRun, TableRow, TableCell, Table, WidthType, PageOrientation, convertInchesToTwip, TableLayoutType } = require("docx");

// Helper: Generate the full DOCX from extracted data
async function generateDocxFile(studentData, outputPath) {
  const studentChunks = chunkArray(studentData, 40);

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: "Arial", size: 22 },
        },
      },
    },
    sections: studentChunks.map((chunk, index) => {
      const firstColumnData = chunk.slice(0, 20);
      const secondColumnData = chunk.slice(20, 40);

      const sideBySideTables = new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [createFixedTable(firstColumnData)],
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                children: [createFixedTable(secondColumnData)],
              }),
            ],
          }),
        ],
        layout: TableLayoutType.FIXED,
        borders: {
          top: { size: 0, color: "FFFFFF" },
          bottom: { size: 0, color: "FFFFFF" },
          left: { size: 0, color: "FFFFFF" },
          right: { size: 0, color: "FFFFFF" },
          insideHorizontal: { size: 0, color: "FFFFFF" },
          insideVertical: { size: 0, color: "FFFFFF" },
        },
      });

      return {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.PORTRAIT,
              width: convertInchesToTwip(8.27),
              height: convertInchesToTwip(11.69),
            },
            margin: {
              top: 1000,
              bottom: 1000,
              left: 900,
              right: 860,
            },
          },
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({ text: `CONFIDENTIAL - Page ${index + 1}`, underline: {}, font: "Arial", size: 18 }),
              new TextRun({ text: "\tDEPARTMENT OF ARCHITECTURE", bold: true, allCaps: true, font: "Arial", size: 22 }),
            ],
            tabStops: [{ type: "center", position: 5050 }],
          }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "OBAFEMI AWOLOWO UNIVERSITY, ILE-IFE", bold: true, allCaps: true, alignment: "center" }),
          new Paragraph({ text: "" }),
          new Paragraph({ text: "EXAMINATION RAW SCORE SHEET", alignment: "center" }),
          new Paragraph({ text: "" }),
          new Paragraph({
            children: [
              new TextRun({ text: "SEMESTER: ..................................." }),
              new TextRun({ text: "\tCOURSE CODE: .........................................................." }),
            ],
            tabStops: [{ type: "right", position: 10500 }],
          }),
          new Paragraph({ text: "" }),
          new Paragraph({
            children: [
              new TextRun({ text: "SESSION: ......................................." }),
              new TextRun({ text: "\tCOURSE TITLE: .........................................................." }),
            ],
            tabStops: [{ type: "right", position: 10500 }],
          }),
          new Paragraph({ text: "" }),
          sideBySideTables,

          new Paragraph({ text: "" }),
          new Paragraph({ text: "" }),
  
          new Paragraph({
            text: "Result Summary",
            allCaps: true,
            font: "Arial",
            size: 22,
          }),
  
          new Paragraph({ text: "" }),
  
          new Paragraph({
            children: [
              new TextRun({ text: "A", bold: true, font: "Arial" }),
              new TextRun({ text: ": 70–100", font: "Arial" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "B", bold: true, font: "Arial" }),
              new TextRun({ text: ": 60–69", font: "Arial" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "C", bold: true, font: "Arial" }),
              new TextRun({ text: ": 50–59", font: "Arial" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "D", bold: true, font: "Arial" }),
              new TextRun({ text: ": 40–49", font: "Arial" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "E", bold: true, font: "Arial" }),
              new TextRun({ text: ": 30–39", font: "Arial" }),
              new TextRun({ text: "\t" }),
              new TextRun({ text: "F", bold: true, font: "Arial" }),
              new TextRun({ text: ": 0–29", font: "Arial" }),
            ],
            tabStops: [
              { type: "left", position: 0 },
              { type: "left", position: 1800 },
              { type: "left", position: 3600 },
              { type: "left", position: 5600 },
              { type: "left", position: 7600 },
              { type: "left", position: 9300 },
            ],
          }),
  
          new Paragraph({ text: "" }),
  
          new Paragraph({ text: "" }),
  
          new Paragraph({ text: "" }),
  
          new Paragraph({
            children: [
              new TextRun({
                text: ".......................................",
                font: "Arial",
                size: 19,
              }),
              new TextRun({
                text: "\t......................................",
                font: "Arial",
                size: 19,
              }),
            ],
            tabStops: [
              {
                type: "right",
                position: 10500, // Adjust this to match the page width (e.g., 9000 for near A4 edge)
              },
            ],
          }),
        ],
      };
    }),
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}


module.exports = { generateDocxFile }
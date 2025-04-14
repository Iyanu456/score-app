const fs = require("fs");
const path = require("path");
const {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  WidthType,
  TableLayoutType,
  PageOrientation,
  convertInchesToTwip,
  Section,
} = require("docx");

// Helper to create one table with fixed rows
const chunkArray = (array, size) => {
  const result = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
};

const createFixedTable = (data = []) => {
  const tableWidth = 5058;
  const columnWidths = [700, 1900, 700, 900, 1000];

  const createCell = (text, isHeader = false, width) => {
    return new TableCell({
      width: { size: width, type: WidthType.DXA },
      margins: {
        top: 100,
        bottom: isHeader ? 200 : 100,
        left: 100,
        right: 100,
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({ text, bold: isHeader, font: "Arial", size: 22 }),
          ],
        }),
      ],
    });
  };

  const headers = ["S/N", "REG. NO.", "C/A", "EXAM", "TOTAL"];
  const headerRow = new TableRow({
    children: headers.map((h, i) => createCell(h, true, columnWidths[i])),
  });

  const dataRows = data.map(
    (student, index) =>
      new TableRow({
        children: [
          createCell(
            (student.sn || index + 1).toString(),
            false,
            columnWidths[0]
          ),
          createCell(student.regNo || "", false, columnWidths[1]),
          createCell(student.ca?.toString() || "", false, columnWidths[2]),
          createCell(student.exam?.toString() || "", false, columnWidths[3]),
          createCell(student.total?.toString() || "", false, columnWidths[4]),
        ],
      })
  );

  const totalRows = 20;
  const blankRowsNeeded = totalRows - dataRows.length;
  const blankRows = Array.from({ length: Math.max(0, blankRowsNeeded) }).map(
    () =>
      new TableRow({
        children: columnWidths.map((width) => createCell("", false, width)),
      })
  );

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    rows: [headerRow, ...dataRows, ...blankRows],
    layout: TableLayoutType.FIXED,
  });
};

// Simulated student data (you can replace this with real data)
const generateRandomData = (numStudents) => {
  const students = [];

  for (let i = 1; i <= numStudents; i++) {
    const regNo = `AB${String(i).padStart(3, "0")}`; // e.g., AB001, AB002
    const ca = Math.floor(Math.random() * 30) + 10; // CA score between 10 and 40
    const exam = Math.floor(Math.random() * 60) + 40; // Exam score between 40 and 100
    const total = ca + exam;

    students.push({ sn: i, regNo, ca, exam, total });
  }

  return students;
};

// Generate data for 50 students
const studentData = generateRandomData(56);

// Split students into chunks of 40 for pagination (20 per column, 2 columns per page)
const studentChunks = chunkArray(studentData, 40);




const createSummaryCell = (text, isHeader) =>
  new TableCell({
    width: { size: 10, type: WidthType.PERCENTAGE },
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            bold: isHeader, // Make header text bold
            
          }),
          
        ],
      }),
    ],
  });

// Header row
const summaryHeaderRow = new TableRow({
  children: [
    createSummaryCell("A", true),
    createSummaryCell("B", true),
    createSummaryCell("C", true),
    createSummaryCell("D", true),
    createSummaryCell("E", true),
    createSummaryCell("F", true),
    createSummaryCell("Total", true),
  ],
});

// Data row
const summaryDataRow = new TableRow({
  children: [
    createSummaryCell(""),
    createSummaryCell(""),
    createSummaryCell(""),
    createSummaryCell(""),
    createSummaryCell(""),
    createSummaryCell(""),
    createSummaryCell(""),
  ],
});

// Create table
const resultSummaryTable = new Table({
  width: {
    size: 60,
    type: WidthType.PERCENTAGE,
  },
  rows: [summaryHeaderRow, summaryDataRow],
  margins: {
    top: 100,
    bottom: 100,
    left: 100,
    right: 100,
  }
});

// Create the document
const doc = new Document({
  styles: {
    default: {
      document: {
        run: {
          font: "Arial",
          size: 22, // 11pt = 22 half-points
        },
      },
    },
  },

  sections: studentChunks.map((chunk, index) => {
    // For each page, create left and right column data
    const firstColumnData = chunk.slice(0, 20);
    const secondColumnData = chunk.slice(20, 40);

    // Create side-by-side tables for this specific chunk
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
            width: convertInchesToTwip(8.27), // A4 width in inches
            height: convertInchesToTwip(11.69), // A4 height in inches
          },
          margin: {
            top: 1000, // 0.5 inch
            bottom: 1000, // 1 inch
            left: 900,
            right: 860,
          },
        },
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: `CONFIDENTIAL - Page ${index + 1}`,
              underline: {},
              font: "Arial",
              size: 18,
            }),
            new TextRun({
              text: "\tDEPARTMENT OF ARCHITECTURE",
              bold: true,
              allCaps: true,
              font: "Arial",
              size: 22,
            }),
          ],
          tabStops: [
            {
              type: "center",
              position: 5050, // Half of page width (approx. center)
            },
          ],
        }),

        new Paragraph({ text: "" }), // Spacer paragraph

        new Paragraph({
          children: [
            new TextRun({
              text: "OBAFEMI AWOLOWO UNIVERSITY, ILE-IFE",
              bold: true,
              size: 24, // 14pt
              font: "Arial",
              allCaps: true,
            }),
          ],
          alignment: "center",
        }),

        new Paragraph({ text: "" }),

        new Paragraph({
          children: [
            new TextRun({
              text: "EXAMINATION RAW SCORE SHEET",
              bold: false,
              size: 20, // 14pt
              font: "Arial",
              allCaps: true,
            }),
          ],
          alignment: "center",
        }),

        new Paragraph({ text: "" }),

        new Paragraph({
          children: [
            new TextRun({
              text: "SEMESTER: ...................................",
              font: "Arial",
              size: 19,
            }),
            new TextRun({
              text: "\tCOURSE CODE: ......................................",
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

        new Paragraph({ text: "" }),

        new Paragraph({
          children: [
            new TextRun({
              text: "SESSION: .......................................",
              font: "Arial",
              size: 19,
            }),
            new TextRun({
              text: "\tCOURSE TITLE: ......................................",
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

        new Paragraph({ text: "" }),

        sideBySideTables, // This table is unique to each chunk of data

        new Paragraph({ text: "" }),

        new Paragraph({
          text: "Result Summary",
          allCaps: true,
          font: "Arial",
          size: 22,
        }),

        new Paragraph({ text: "" }),

        resultSummaryTable, // This table is unique to each chunk of data

        new Paragraph({ text: "" }),

        /*new Paragraph({
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
        }),*/

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

// Save the document
const outputPath = path.join(__dirname, "scores.docx");

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outputPath, buffer);
  console.log("✅ DOCX file created at:", outputPath);
});

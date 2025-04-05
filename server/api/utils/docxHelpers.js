const { Packer } = require("docx");
const fs = require("fs");
const path = require("path");
const { Document, Paragraph, TextRun, TableRow, TableCell, Table, WidthType, PageOrientation, convertInchesToTwip, TableLayoutType } = require("docx");

const chunkArray = (array, size) => {
    const result = [];
    for (let i = 0; i < array.length; i += size) {
      result.push(array.slice(i, i + size));
    }
    return result;
  };




const createFixedTable = (data = []) => {
  const tableWidth = 5058;
  const columnWidths = [800, 1900, 600, 900, 1000];

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


module.exports = {
  chunkArray,
  createFixedTable,
};
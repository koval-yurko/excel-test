import { Workbook, Worksheet } from 'exceljs';

const data = [
    ['one', 'two', 'three'],
    [1, 2, 3],
    [4, 5, 6],
];

const wb = new Workbook();
const ws = wb.addWorksheet(
    "test", { views: [{ state: 'frozen', xSplit: 1, ySplit: 1 }] },
);

ws.addRows(data);

ws.getCell(1, 1).fill = {
    type:"pattern",
    pattern: "solid",
    bgColor: { argb: "FF00FF00" },
    fgColor: { argb: "00FF00FF" }
};

ws.getCell(2, 2).fill = {
    type:"pattern",
    pattern: "solid",
    bgColor: { argb: "FF00FF00" },
    fgColor: { argb: "00FF00FF" }
};

ws.getCell(3, 3).fill = {
    type:"pattern",
    pattern: "solid",
    bgColor: { argb: "FF00FF00" },
    fgColor: { argb: "00FF00FF" }
};

ws.getCell(3, 1).border = {
    top: { style: "thin", color: { argb: "00FF00FF" } }
};

ws.mergeCells([1, 2, 1, 3]);

wb.xlsx.writeFile('test2.xlsx');

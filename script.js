const file = require("fs");
const readXlsxFile = require("read-excel-file/node");
const xlnode = require("excel4node");

const Symbols = "abcdefghijklmnopqrstuvwxyz";

let mapData = file.readFileSync(
  "/home/anjaliakansha/Desktop/node_automation/data/mapping.json"
);

mapData = JSON.parse(mapData);
let dataMap = new Map();

readXlsxFile(
  "/home/anjaliakansha/Desktop/node_automation/data/Sample-Template-File.xlsx"
).then((rows) => {
  for (let i = 0; i < rows.length; i++) {
    for (let j = 0; j < rows[0].length; j++) {
      dataMap.set(rows[i][j], [i + 1, j + 1]);
    }
  }

  readXlsxFile(
    "/home/anjaliakansha/Desktop/node_automation/data/Sample-Data-File.xlsx"
  ).then((rows) => {
    for (let i = 1; i < rows.length; i++) {
      let workbook = new xlnode.Workbook();
      let worksheet = workbook.addWorksheet("studentData");

      for (let column of dataMap.keys()) {
        worksheet
          .cell(dataMap.get(column)[0], dataMap.get(column)[1])
          .string(column);
      }

      for (let j = 0; j < rows[i].length; j++) {
        var cell = mapData[rows[0][j].toLowerCase()]
          .replace(/\'/g, "")
          .split(/(\d+)/)
          .filter(Boolean);
        var column = cell[0];
        var row = cell[1];

        var colIndex = Symbols.search(column);
        worksheet.cell(row, colIndex + 1).string(rows[i][j].toString());
      }

      workbook.write(`Output/mappedData${i}.xlsx`);
    }
  });
});

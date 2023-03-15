const XLSX = require('xlsx');
const _ = require('lodash');
const express = require('express');
const multer = require('multer');
const cors = require('cors')

const app = express();
app.use(cors());
const upload = multer();


function excelSheetsCompare(sheet1, sheet2) {
  const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
  const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });

  const rowDifferences = {};
  const maxRow = Math.max(data1.length, data2.length);
  const maxCol = Math.max(data1[0].length, data2[0].length);

  const headers = data1[0]; // use the first row of sheet1 as headers
  for (let i = 1; i < maxRow; i++) { // start from the second row
    const rowDiffs = []; // create an array for the differences in this row
    for (let j = 0; j < maxCol; j++) {
      const val1 = data1[i] && data1[i][j] !== undefined ? data1[i][j] : '';
      const val2 = data2[i] && data2[i][j] !== undefined ? data2[i][j] : '';
      rowDiffs.push({
        col: headers[j],
        val1,
        val2
      });
    }
    if (rowDiffs.length > 0) {
      rowDifferences[i] = rowDiffs; // add the row's differences to the object
    }
  }

  return rowDifferences;
}

app.post('/compare-excel', upload.fields([{ name: 'file1' }, { name: 'file2' }]), (req, res) => {
  const file1 = req.files.file1[0];
  const file2 = req.files.file2[0];

  const workbook1 = XLSX.read(file1.buffer, { type: 'buffer' });
  const workbook2 = XLSX.read(file2.buffer, { type: 'buffer' });

  const sheetNames1 = workbook1.SheetNames;
  const excelDataDifference = [];

  for (const sheetName of sheetNames1) {
    const sheet1 = workbook1.Sheets[sheetName];
    const sheet2 = workbook2.Sheets[sheetName];

    const rowDifferences = excelSheetsCompare(sheet1, sheet2);
    if (Object.keys(rowDifferences).length > 0) {
      excelDataDifference.push({
        rowDifferences
      });
    }
  }

  res.json(excelDataDifference);
});

app.listen(3000, () => {
  console.log('Server started on port 3000');
});

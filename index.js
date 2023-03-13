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

  if (_.isEqual(data1, data2)) {
    return [];
  }

  const sheetdifferences = [];

  for (let i = 0; i < data1.length; i++) {
    for (let j = 0; j < data1[i].length; j++) {
      if (data1[i][j] !== data2[i][j]) {
        const colName = XLSX.utils.encode_col(j);
        sheetdifferences.push({
          row: i + 1,
          col: data1[0][j], 
          val1: data1[i][j],
          val2: data2[i][j]
        });
      }
    }
  }
  
  return sheetdifferences;
}

app.post('/compare-excel', upload.fields([{ name: 'file1' }, { name: 'file2' }]), (req, res) => {
  const file1 = req.files.file1[0];
  const file2 = req.files.file2[0];

  const workbook1 = XLSX.read(file1.buffer, { type: 'buffer' });
  const workbook2 = XLSX.read(file2.buffer, { type: 'buffer' });

  const sheetNames1 = workbook1.SheetNames;
 
  const excelDataDifference = [];

  for (const excelfileName of sheetNames1) {
    const excelsheet1 = workbook1.Sheets[excelfileName];
    const excelsheet2 = workbook2.Sheets[excelfileName];

    const excelDifference = excelSheetsCompare(excelsheet1, excelsheet2);
    if (excelDifference && excelDifference.length > 0) {
      excelDataDifference.push({
        excelfileName,
        excelDifference
      });
    }
  }

  const result = {
    excelDataDifference
  };

  res.json(result);
});

app.listen(3000, () => {
  console.log('Server started on port 3000');
});
const XLSX = require('xlsx');

// Function to read Excel file and extract data
function readExcelFile(file) {
  const workbook = XLSX.readFile(file);
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
}

// Function to filter and copy data to another sheet
function getProdHour1(data) {
  const Sheet1 = data.Sheet1;
  const Sheet2 = data.Sheet2;

  const finalrow = Sheet2['!ref'].split(':')[1].replace(/\D/g, '');
  const columnCount = 35;

  for (let i = 1; i <= columnCount; i++) {
    if (Sheet2[String.fromCharCode(64 + i) + '1'] === 'Presort Ivc') {
      const range = Sheet2[String.fromCharCode(64 + i) + '4:' + String.fromCharCode(64 + i) + finalrow];
      Sheet1[String.fromCharCode(66) + '1'].v += 1;
      Sheet1[String.fromCharCode(66) + (Sheet1[String.fromCharCode(66) + '1'].v)].v = range.v;
    }
    if (Sheet2[String.fromCharCode(64 + i) + '1'] === 'Presort Vcp') {
      const range = Sheet2[String.fromCharCode(64 + i) + '4:' + String.fromCharCode(64 + i) + finalrow];
      Sheet1[String.fromCharCode(68) + '1'].v += 1;
      Sheet1[String.fromCharCode(68) + (Sheet1[String.fromCharCode(68) + '1'].v)].v = range.v;
    }
    if (Sheet2[String.fromCharCode(64 + i) + '1'] === 'Presort Fulfillment Vcp') {
      const range = Sheet2[String.fromCharCode(64 + i) + '4:' + String.fromCharCode(64 + i) + finalrow];
      Sheet1[String.fromCharCode(67) + '1'].v += 1;
      Sheet1[String.fromCharCode(67) + (Sheet1[String.fromCharCode(67) + '1'].v)].v = range.v;
    }
    if (Sheet2[String.fromCharCode(64 + i) + '1'] === 'Presort Udc') {
      const range = Sheet2[String.fromCharCode(64 + i) + '4:' + String.fromCharCode(64 + i) + finalrow];
      Sheet1[String.fromCharCode(69) + '1'].v += 1;
      Sheet1[String.fromCharCode(69) + (Sheet1[String.fromCharCode(69) + '1'].v)].v = range.v;
    }
  }
}

// Function to find names and copy to another sheet
function findNamesHour1(data) {
  const Sheet1 = data.Sheet1;
  const Sheet2 = data.Sheet2;

  const PullSheet = Sheet2['A4:A300'];

  for (let i = 0; i < PullSheet.length; i++) {
    const Name = PullSheet[i][0];
    const targetSheet = Name > 0 ? Sheet1 : Sheet2;

    const PasteCell = targetSheet['A2'].v === '' ? targetSheet['A2'] : targetSheet['A1'].v + 1;
    targetSheet[String.fromCharCode(65) + PasteCell].v = Name;
  }
}

// Function to transfer ZNum to Name
function transferZNumtoNameHour1(data) {
  const Sheet1 = data.Sheet1;
  const Sheet27 = data.Sheet27;

  const Znum = Sheet1['A2:A400'];
  const Names = Sheet27['B2:B4000'];

  for (let i = 0; i < Names.length; i++) {
    const nameRange = Names[i];
    const zNumRange = Znum[i];

    Znum['!ref'] = Znum['!ref'].replace('A2:A400', 'A2:A' + Znum.length);

    Znum[String.fromCharCode(65) + zNumRange.r].v = Znum[String.fromCharCode(65) + zNumRange.r].v.replace(
      nameRange.v,
      nameRange.v + ' ' + nameRange.c
    );
  }
}

// Usage example
const excelData = readExcelFile('data.xlsx');
getProdHour1(excelData);
findNamesHour1(excelData);
transferZNumtoNameHour1(excelData);

const XLSX = require('xlsx');

// Function to read Excel file and extract data
function readExcelFile(file) {
  const workbook = XLSX.readFile(file);
  const worksheet = workbook.Sheets['Sheet2'];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
}

// Function to generate table data
function generateTableData(data) {
  const tableData = [];

  // Generate table header
  const header = data[0];
  header.push('Total');

  tableData.push(header);

  // Generate table body rows
  for (let i = 3; i < data.length; i++) {
    const row = data[i];

    // Extract values from columns A, F, and G
    const teamMember = row[0];
    const presortIvc = row[5];
    const presortVcp = row[6];

    const total = presortIvc + presortVcp;

    // Add total to the row
    row.push(total);

    tableData.push(row);
  }
  return tableData;
}

// Function to generate HTML table content
function generateHTMLTableContent(tableData, tableId) {
  let html = `<table id="${tableId}">\n`;

  // Generate table header row
  html += '<thead>\n';
  html += '<tr>\n';
  tableData[0].forEach((column) => {
    html += `<th>${column}</th>\n`;
  });
  html += '</tr>\n';
  html += '</thead>\n';

  // Generate table body rows
  html += '<tbody>\n';
  for (let i = 1; i < tableData.length; i++) {
    const row = tableData[i];
    html += '<tr>\n';
    row.forEach((cell) => {
      html += `<td>${cell}</td>\n`;
    });
    html += '</tr>\n';
  }
  html += '</tbody>\n';

  html += '</table>';

  return html;
}

module.exports = {
  readExcelFile,
  generateTableData,
  generateHTMLTableContent,
};

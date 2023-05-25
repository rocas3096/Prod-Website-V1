const fs = require('fs');
const excelParser = require('../functions/excelParser');

function generateParsedTable() {
  const excelData = excelParser.readExcelFile('Prod Tracker Updated 5-25.xlsm');

  // Perform filtering or calculations specific to parsed table
  // Example: Only include rows with non-empty values in "Presort Ivc" or "Presort Vcp"
  const filteredData = excelData.filter((row, index) => index > 2 && (row[5] || row[6]));

  // Generate parsed table data based on the filtered data
  const parsedTableData = filteredData.map((row) => {
    const teamMember = row[0];
    const presortIvc = row[5];
    const presortVcp = row[6];
    const total = presortIvc + presortVcp;

    return [teamMember, presortIvc, presortVcp, total];
  });

  // Sort the parsed table data in descending order based on the "Total" column
  parsedTableData.sort((a, b) => b[3] - a[3]);

  // Generate custom HTML table content
  let html = '<table id="parsed-table">\n';

  // Generate table header row
  html += '<thead>\n';
  html += '<tr>\n';
  html += '<th>Team Members</th>\n';
  html += '<th>Presort IVC</th>\n';
  html += '<th>Presort VCP</th>\n';
  html += '<th>Total</th>\n';
  html += '</tr>\n';
  html += '</thead>\n';

  // Generate table body rows
  html += '<tbody>\n';
  parsedTableData.forEach((row) => {
    html += '<tr>\n';
    row.forEach((cell) => {
      html += `<td>${cell}</td>\n`;
    });
    html += '</tr>\n';
  });
  html += '</tbody>\n';

  html += '</table>';

  // Generate navigation bar HTML
  const navigationBar = `
    <div class="navigation-bar">
      <a href="../index.html" class="nav-link">Home</a>
      <a href="Data.html" class="nav-link">Data</a>
    </div>
  `;

  // Combine navigation bar and parsed table content
  const parsedHtmlContent = navigationBar + html;

  const parsedHtmlFilePath = 'pages/ParsedData.html';
  fs.writeFileSync(parsedHtmlFilePath, parsedHtmlContent);

  console.log(`Parsed data table written to ${parsedHtmlFilePath}`);
}

module.exports = {
  generateParsedTable,
};

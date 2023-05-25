const fs = require('fs');
const excelParser = require('../functions/excelParser');

function generateDataTable() {
  const excelData = excelParser.readExcelFile('Prod Tracker Updated 5-25.xlsm');

  // Perform filtering or calculations specific to data table
  // Example: Include all rows and calculate the total sum of "Presort Ivc" and "Presort Vcp"
  const dataTableData = excelParser.generateTableData(excelData);
  const dataTableContent = excelParser.generateHTMLTableContent(dataTableData, 'excel-table');

  // Generate navigation bar HTML
  const navigationBar = `
    <div class="navigation-bar">
      <a href="../index.html" class="nav-link">Home</a>
      <a href="ParsedData.html" class="nav-link">Filter</a>
    </div>
  `;

  // Combine navigation bar and data table content
  const htmlContent = navigationBar + dataTableContent;

  const htmlFilePath = 'pages/Data.html';
  fs.writeFileSync(htmlFilePath, htmlContent);

  console.log(`Data table written to ${htmlFilePath}`);
}

module.exports = {
  generateDataTable,
};

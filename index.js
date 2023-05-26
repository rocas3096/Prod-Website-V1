const parsedTableGenerator = require("./javaPages/parsedTableGenerator");
const dataTableGenerator = require("./javaPages/dataTableGenerator");

function generateTables() {
  parsedTableGenerator.generateParsedTable();
  dataTableGenerator.generateDataTable();
}

generateTables();

const XLSX = require("xlsx");
const fs = require("fs");

const jsonData = JSON.parse(fs.readFileSync("output2.json", "utf-8"));

const formattedData = [];

jsonData.forEach((row) => {
  const formattedRow = {
    clientName: row.clientName
  };

  row.results.forEach((item) => {
    const service = item.service;

    if (service === "perm") {
      formattedRow.PermTerms = item.terms;
      formattedRow.Permscript = item.script;
      formattedRow.Permoutput = item.output ? JSON.stringify(item.output)
  : "";
    }

    
  });

  formattedData.push(formattedRow);
});

const worksheet = XLSX.utils.json_to_sheet(formattedData);
const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
const fileName = "C:/projects/JsonToExcel/perm2Xcel.xlsx";
XLSX.writeFile(workbook, fileName);

console.log(`Excel file created: ${fileName}`);
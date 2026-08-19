const XLSX = require("xlsx");
const fs = require("fs");

const jsonData = JSON.parse(fs.readFileSync("output.json", "utf-8"));

const formattedData = [];

jsonData.forEach((row) => {
  const formattedRow = {
    ClientName: row.clientName
  };

  row.results.forEach((item) => {
    const service = item.service;

    if (service === "Payroll") {
      formattedRow.PayrollTerms = item.terms;
      formattedRow.PayrollScript = item.script;
      formattedRow.PayrollOutput = item.output ? JSON.stringify(item.output)
  : "";
    }

    if (service === "MidMonth") {
      formattedRow.MidMonthTerms = item.terms;
      formattedRow.MidMonthScript = item.script;
      formattedRow.MidMonthOutput = item.output ? JSON.stringify(item.output)
  : "";
    }

    if (service === "Reimbursement") {
      formattedRow.ReimbursementTerms = item.terms;
      formattedRow.ReimbursementScript = item.script;
      formattedRow.ReimbursementOutput = item.output ? JSON.stringify(item.output)
  : "";
    }

    if (service === "SourcingFee") {
      formattedRow.SourcingFee = item.terms;
      formattedRow.SourcingFeeScript = item.script;
      formattedRow.SourcingFeeOutput = item.output ? JSON.stringify(item.output)
  : "";
    }

    if (service === "AbsorptionFee") {
      formattedRow.AbsorptionFeePayroll = item.terms;
      formattedRow.AbsorptionFeeScript = item.script;
      formattedRow.AbsorptionFeeOutput = item.output ? JSON.stringify(item.output)
  : "";
    }
  });

  formattedData.push(formattedRow);
});

const worksheet = XLSX.utils.json_to_sheet(formattedData);
const workbook = XLSX.utils.book_new();

XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
const fileName = "C:/projects/JsonToExcel/staffingXcel.xlsx";
XLSX.writeFile(workbook, fileName);

console.log(`Excel file created: ${fileName}`);
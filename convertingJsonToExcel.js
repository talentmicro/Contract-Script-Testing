// const XLSX = require('xlsx');

// const fs = require('fs');
// const jsonData = JSON.parse(fs.readFileSync('data.json', 'utf-8'));
// jsonData.forEach(item => {
//     j = JSON.parse(item.claimDetails);
//     console.log(j)
// });

// const excelHeaders = [
//     { header: "Emp Code", key: "employeeCode" },
//     { header: "Name", key: "senderName" },
//     { header: "WM Ref.", key: "parentId" },
//     { header: "Ref. No.", key: "referenceNumber" },
//     { header: "Cost Center", key: "costCenterCode" },
//     { header: "GL Code", key: "j?.glCode" },
//     { header: "Expense Type", key: "j?.expType" },
//     { header: "Expense Date", key: "j?.expDate" },
//     { header: "Particular", key: "j?.particulars" },
//     { header: "Amount", key: "totalAmount" },
//     { header: "Rejected Amount", key: "" },
//     { header: "Submitted On", key: "createdDate" },
//     { header: "Status", key: "statusTitle" },
//     { header: "Approved On", key: "approvedOn" },
//     { header: "Settled", key: "settled" },
//     { header: "Processed On", key: "processedOn" },
//     { header: "Approver Name", key: "approver" },
//     { header: "Reciever Name", key: "receiverName" },
//     { header: "Reciever Notes", key: "receiverNotes" },
// ];

// const formattedData = jsonData.map(row => {
//     const formattedRow = {};
//     excelHeaders.forEach(col => {
//         formattedRow[col.header] = row[col.key];
//     });
//     console.log(formattedRow)
//     return formattedRow;
// });

// const workbook = XLSX.utils.book_new();
// const worksheet = XLSX.utils.json_to_sheet(formattedData);

// XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

// const fileName = "C:/Users/HireCraft/Documents/DataFromJsonToExcel.xlsx"
// XLSX.writeFile(workbook, fileName);

// console.log(`Excel file created: ${fileName}`);

const XLSX = require("xlsx");
const fs = require("fs");

const jsonData = JSON.parse(fs.readFileSync("data.json", "utf-8"));

const excelHeaders = [
  { header: "Emp Code", key: "employeeCode" },
  { header: "Name", key: "senderName" },
  { header: "WM Ref.", key: "parentId" },
  { header: "Ref. No.", key: "referenceNumber" },
  { header: "Cost Center", key: "costCenterCode" },
  { header: "GL Code", key: "GLCode" },
  { header: "Expense Type", key: "expType" },
  { header: "Expense Date", key: "expDate" },
  { header: "Particular", key: "particulars" },
  { header: "Amount", key: "amount" },
  { header: "Rejected Amount", key: "" },
  { header: "Submitted On", key: "createdDate" },
  { header: "Status", key: "statusTitle" },
  { header: "Approved On", key: "approvedOn" },
  { header: "Settled", key: "settled" },
  { header: "Processed On", key: "processedOn" },
  { header: "Approver Name", key: "approver" },
  { header: "Receiver Name", key: "receiverName" },
  { header: "Receiver Notes", key: "receiverNotes" },
];

const formattedData = [];

jsonData.forEach((row) => {
  try {
    const claimDetails = JSON.parse(row.claimDetails); // Parse claimDetails array

    claimDetails.forEach((detail) => {
      const formattedRow = {};

      excelHeaders.forEach((col) => {
        formattedRow[col.header] = detail[col.key] || row[col.key] || ""; // Use detail if available, else row
      });

      formattedData.push(formattedRow);
    });
  } catch (error) {
    console.error("Error parsing claimDetails for row:", row, error);
  }
});

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(formattedData);

XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

const fileName = "C:/Users/HireCraft/Documents/DataFromJsonToExcel.xlsx";
XLSX.writeFile(workbook, fileName);

console.log(`Excel file created: ${fileName}`);

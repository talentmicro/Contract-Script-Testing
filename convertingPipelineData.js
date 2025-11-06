const XLSX = require("xlsx");

const fs = require("fs");
// const jsonData = JSON.parse(fs.readFileSync('data.json', 'utf-8'));
const jsonData = JSON.parse(
  fs.readFileSync("quess-offerpipeline-all-data.json", "utf-8")
);
// jsonData.forEach(item => {
//     claimDetails = JSON.parse(item.claimDetails);
//     console.log(claimDetails)
// });

// const excelHeaders = [
//     { header: "Emp Code", key: "employeeCode" },
//     { header: "Name", key: "senderName" },
//     { header: "WM Ref.", key: "parentId" },
//     { header: "Ref. No.", key: "referenceNumber" },
//     { header: "Cost Center", key: "costCenterCode" },
//     { header: "GL Code", key: "glCode" },
//     { header: "Expense Type", key: "expType" },
//     { header: "Expense Date", key: "expDate" },
//     { header: "Particular", key: "particulars" },
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

const excelHeaders = [
  { header: "SLNO", key: "SLNO" },
  { header: "Client/ Department", key: "Requirement_Jobtitle" },
  { header: "Client Code", key: "ClientCode" },
  { header: "Recruiter ID", key: "EmpCode" },
  { header: "Job Title", key: "Requirement_Jobtitle" },
  { header: "Aadhar Number]", key: "Applicant_Aadhar_Number" },
  { header: "Job Code]", key: "Requirement_Jobcode" },
  { header: "Mandate Created Date", key: "Requirement_Start_Date" },
  { header: "Candidate Name", key: "Applicant_Full_Name" },
  { header: "Contact number", key: "Applicant_Mobile_Number" },
  { header: "Email id", key: "Applicant_Email_ID" },
  { header: "Current company", key: "Applicant_Present_Employer" },
  { header: "Total experience]", key: "Applicant_Experience" },
  { header: "Notice period", key: "Applicant_Notice_Period" },
  { header: "Current ctc", key: "Applicant_Present_Salary" },
  { header: "Offered Ctc", key: "Offer_CTC" },
  { header: "Designation Offered", key: "DesignationTitle" },
  { header: "Current location", key: "Applicant_Present_Location" },
  { header: "Offered Location", key: "OfferedLocationTitle" },
  { header: "Joining Date", key: "Actual_Joining_Date" },
  { header: "Stage_Status", key: "Applicant_Stage_Status" },
  { header: "Recruiter", key: "Recruiter" },
  { header: "Sourcer", key: "sourcer" },
  { header: "Notes", key: "Notes" },
  { header: "State", key: "state" },
  { header: "Offered Date", key: "offered_date" },
  { header: "Planned Date", key: "planned_date" },
  { header: "Offer Expiry Date", key: "offer_expiry_date" },
  { header: "Contract Start Date", key: "contract_start_date" },
  { header: "Contract End Date", key: "contract_end_date" },
  { header: "BU Entity", key: "BU_entity" },
  { header: "IVR Type", key: "IVR_Type" },
  { header: "Employement Type", key: "employment_type" },
  { header: "Source Type", key: "source_type" },
  { header: "integration Response", key: "" },
  { header: "Integration Offer Id", key: "integrationOfferId" },
  { header: "Integration Record Id", key: "integrationRecordId" },
  { header: "Organizational Unit", key: "Organizational_Unit" },
  { header: "Profile Type", key: "Profile_Type" },
  { header: "Sales Office", key: "Sales_Office" },
  { header: "User Type", key: "User_Type" },
  { header: "OMS Division", key: "OMS_Division" },
  { header: "Reporting Manager Name", key: "Reporting_Manager_Name" },
  { header: "Department Name", key: "Department_Name" },
  { header: "Reporting Manager Email", key: "Reporting_Manager_Email" },
  { header: "Reporting Manager ID", key: "Reporting_Manager_ID" },
  { header: "Sales Office Pincode", key: "Sales_Office_Pincode" },
];

const formattedData = jsonData.map((row) => {
  const formattedRow = {};
  excelHeaders.forEach((col) => {
    formattedRow[col.header] = row[col.key];
  });
  return formattedRow;
});

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(formattedData);

XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

// const fileName = "C:/Users/HireCraft/Documents/DataFromJsonToExcel.xlsx"
const fileName = "output.xlsx";
XLSX.writeFile(workbook, fileName);

console.log(`Excel file created: ${fileName}`);

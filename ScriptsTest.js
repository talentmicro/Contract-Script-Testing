const XLSX = require("xlsx");
const fs = require("fs");
 
const workbook = XLSX.readFile("input.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
 
const rows = XLSX.utils.sheet_to_json(sheet, {
  header: "A",
  defval: null,
  raw: true
});
 
const COL = {
  client: "C",
  payrollTerms: "Q",
  payroll: "R",
  midmonthTerms: "T",
  midmonth: "U",
  reimbursementTerms: "W",
  reimbursement: "X",
  sourcingTerms: "AB",
  sourcing: "AC",
  replacementPeriod: "AD",
  absorptionTerms: "AF",
  absorption: "AG"
};
 
const baseData = {
  contract: {
    Input: "mpg",
    startDate: "2025-11-18",
    endDate: "2026-11-18"
  },
 BillableCTC: 1200000,
  associate: {
    monthlyCTC: 30000,
    pws: 0,
    TotalWorkedDays:30,
    processedAmount: 10000,
    BillableCTC: 1200000,
    finalCTC: 100000,
    experience: 5,
    DOJ: "2026-01-30",
    DOJ2: "2027-01-23"
  },
 
  taglist: {
    HeadCount: 76,
    interviewSelectedDate: "2026-04-10",
    jobCreatedDate: "2026-01-15",
    jobPostedDate: "2026-03-05",
    replacementPeriod: 30,
    tenureInDays: 365,
    DOJ: "2026-04-20",
    source:"eereferred by client"
  },

  associatetaglist: {
    HeadCount: 120,
    interviewSelectedDate: "2026-04-10",
    jobCreatedDate: "2026-03-01",
    jobPostedDate: "2026-03-05",
    replacementPeriod: 30,
    tenureInDays: 365,
    DOJ: "2026-04-20"
  }
};
 
function clean(script) {
  if (!script || typeof script !== "string") return null;
 
  return script
    .split("//")[0]
    .replace(/\{associate\./g, "associate.")
    .replace(/\{contract\./g, "contract.")
    .replace(/\{taglist\./g, "taglist.")
    .replace(/\{associatetaglist\./g, "associatetaglist.")
    .replace(/\{Input\}/g, "contract.Input")
    .replace(/\{BillableCTC\}/g, "associate.BillableCTC")
    .replace(/\}/g, "")
    // .replace(/^\s*\w+\s*=\s*/i, "")
    .trim();
}
 
// ======================
// RUN SCRIPT
// ======================
// function run(script, ctx) {
//   try {
//     const c = clean(script);
//     if (!c) return null;
 
//     const contract = ctx.contract;
//     const associate = ctx.associate;
//     const taglist = ctx.taglist;
//     const associatetaglist = ctx.associatetaglist;
 
//     const Input = contract?.Input;
//     console.log(Input)
//     return eval(c);
//   } catch (e) {
//     return `Error: ${e.message}`;
//   }
// }
 function extractVariables(script) {
  if (!script) return [];

  const matches = script.match(/\b([a-zA-Z_]\w*)\s*=/g);

  return matches
    ? matches.map(v => v.replace("=", "").trim())
    : [];
}

function run(script, ctx) {
  try {
    const c = clean(script);
    if (!c) return null;

    const contract = ctx.contract;
    const associate = ctx.associate;
    const taglist = ctx.taglist;
    const associatetaglist = ctx.associatetaglist;

    const Input = contract?.Input;

    const vars = extractVariables(c);

   const wrapped = `
  var MarkUp, ReplacementPeriod, ExecutedServiceTerms,Totalpay;

  ${c}

  const __result = {};
  if (typeof MarkUp !== "undefined") __result["MarkUp"] = MarkUp;
  if (typeof ReplacementPeriod !== "undefined") __result["ReplacementPeriod"] = ReplacementPeriod;
  if (typeof ExecutedServiceTerms !== "undefined") __result["ExecutedServiceTerms"] = ExecutedServiceTerms;
  if (typeof Totalpay !== "undefined") __result["Totalpay"] = Totalpay;

  __result;
`;

    return eval(wrapped);

  } catch (e) {
    return `Error: ${e.message}`;
  }
}
// ======================
// EXTRACT INPUTS (SAFE)
// ======================
function extractInputs(script) {
  const cleaned = clean(script);
  const matches = cleaned.match(/\b(mpg|tc|olda|newa|pws|no-pws|contractualstaffing|cr|i|ni|blu|wht|bea|bec|casm|ipod|fp|o|p1|p2|e&b|ae|admin|0-1y|1.1-2y|2.1-3y|3.1-4y|4.1-5y|5.1-6y|6.1-7y|7.1-8y|8.1-9y|9.1-10y|10.1-11y|11.1-12y|12.1-13y|13.1-14y|14.1-15y|15.1-16y|16.1-17y|17.1-18y|18.1-19y|19.1-20y|>20y|jm_administrative_band|jm_operational_band|mm_professional_band|sm_tactical_band|exe_band_and_strategic_band|atr_finance_band3|atr_finance_band4|atr_finance_band5|atr_finance_band6|atr_finance_band7|atr_finance_band8|atr_finance_band9|eaa_generalist_band3|eaa_generalist_band4|eaa_generalist_band5|eaa_generalist_ban6|eaa_generalist_band7|eaa_generalist_band8|eaa_generalist_band9|faba_generalist_band3|faba_generalist_band4|faba_generalist_band5|faba_generalist_band6|faba_generalist_band7|faba_generalist_band8|faba_generalist_band9|hrp_generalist_band3|hrp_generalist_band4|hrp_generalist_band5|hrp_generalist_band6|hrp_generalist_band7|hrp_generalist_band8|hrp_generalist_band9|pds_lead_to_cash_band3|pds_lead_to_cash_band4|pds_lead_to_cash_band5|pds_lead_to_cash_band6|pds_lead_to_cash_band7|pds_lead_to_cash_band8|pds_lead_to_cash_band8|pop_generalist_band3|pop_generalist_band4|pop_generalist_band5|pop_generalist_band6|pop_generalist_band7|pop_generalist_band8|pop_generalist_band9|rsp_generalist_band3|rsp_generalist_band4|rsp_generalist_band5|rsp_generalist_band6|rsp_generalist_band7|rsp_generalist_band8|atr-finance-b3|atr-finance-b4|atr-finance-b5|atr-finance-b6|atr-finance-b7|atr-finance-b8|atr-finance-b9|eaa-generalist-b3|eaa-generalist-b4|eaa-generalist-b5|eaa-generalist-b6|eaa-generalist-b7|eaa-generalist-b8|eaa-generalist-b9|faba-generalist-b3|faba-generalist-b4|faba-generalist-b5|faba-generalist-b6|faba-generalist-b7|faba-generalist-b8|faba-generalist-b9|hrp-generalist-b3|hrp-generalist-b4|hrp-generalist-b5|hrp-generalist-b6|hrp-generalist-b7|hrp-generalist-b8|hrp-generalist-b9|pds-lead-to-cash-b3|pds-lead-to-cash-b4|pds-lead-to-cash-b5|pds-lead-to-cash-b6|pds-lead-to-cash-b7|pds-lead-to-cash-b8|pds-lead-to-cash-b9|pop-generalist-b3|pop-generalist-b4|pop-generalist-b5|pop-generalist-b6|pop-generalist-b7|pop-generalist-b8|pop-generalist-b9|rsp-generalist-b3|rsp-generalist-b4|rsp-generalist-b5|rsp-generalist-b6|rsp-generalist-b7|rsp-generalist-b8|rsp-generalist-b9)\b/g);
  return matches ? [...new Set(matches)] : [];
}
 
// ======================
// GET CELL
// ======================
function get(row, colLetter) {
  return row[colLetter] || null;
}
 
// ======================
// MAIN
// ======================
const output = [];
 
rows.forEach((row, index) => {
  if (index === 0) return;
 
  const excelRow = index + 1;
  const clientName = get(row, COL.client) || "UNKNOWN_CLIENT";
 
  const rules = [
    { name: "Payroll", script: get(row, COL.payroll), terms: get(row, COL.payrollTerms) },
    { name: "MidMonth", script: get(row, COL.midmonth), terms: get(row, COL.midmonthTerms) },
    { name: "Reimbursement", script: get(row, COL.reimbursement), terms: get(row, COL.reimbursementTerms) },
    { name: "SourcingFee", script: get(row, COL.sourcing), terms: get(row, COL.sourcingTerms),replacementPeriod: get(row, COL.replacementPeriod) },
    { name: "AbsorptionFee", script: get(row, COL.absorption), terms: get(row, COL.absorptionTerms) },
  ];
 
  const results = [];
 
  rules.forEach(r => {
    if (!r.script) return;
 
    const inputs = extractInputs(r.script);
 
    let result;
 
    // ======================
    // MULTI INPUT CASE
    // ======================
    if (inputs.length > 0) {
      result = {};
 
      inputs.forEach(val => {
        const ctx = {
          ...baseData,
          contract: {
            ...baseData.contract,
            Input: val
          }
        };
 
        result[val] = run(r.script, ctx);
      });
    }
 
    // ======================
    // SINGLE CASE
    // ======================
    else {
      result = run(r.script, baseData);
    }
 
    results.push({
      service: r.name,
      terms: r.terms,
      replacementPeriod: r.replacementPeriod,
      script: r.script,
      output: result
    });
  });
 
  output.push({
    clientName,
    excelRow,
    results
  });
});
 
// ======================
// WRITE OUTPUT
// ======================
fs.writeFileSync("output.json", JSON.stringify(output, null, 2));
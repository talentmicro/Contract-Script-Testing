const XLSX = require("xlsx");
const fs = require("fs");
 
// const workbook = XLSX.readFile("goliveclient.xlsx");
// const workbook = XLSX.readFile("ServiceTerms_Perm_Clients.xlsx");
const workbook = XLSX.readFile("perm1.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
 
const rows = XLSX.utils.sheet_to_json(sheet, {
  header: "A",
  defval: null,
  raw: true
});
 
const COL = {
  client: "A",
  payrollTerms: "B",
  payroll: "C",
  replacementPeriod: "D"
};
// const COL = {
//   client: "B",
//   payrollTerms: "C",
//   payroll: "D",
//   replacementPeriod: "E"
// };
 
const baseData = {
  contract: {
    Input: "mpg",
    startDate: "2025-04-25",
    endDate: "2027-04-25"
  },
 BillableCTC: 800000,
  associate: {
    monthlyCTC: 50000,
    processedAmount: 10000,
    BillableCTC: 800000,
    finalCTC: 800000,
    experience: 5,
    DOJ: "2026-04-20"
  },
 
  taglist: {
    HeadCount: 120,
    interviewSelectedDate: "2026-04-10",
    jobCreatedDate: "2026-03-01",
    jobPostedDate: "2026-03-05",
    replacementPeriod: 30,
    tenureInDays: 40,
    DOJ: "2026-04-20"
  },
   associatetaglist: {
    HeadCount: 120,
    source:"referred by client",
    interviewSelectedDate: "2026-04-10",
    jobCreatedDate: "2026-03-01",
    jobPostedDate: "2026-03-05",
    replacementPeriod: 30,
    tenureInDays: 40,
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

    const Input = (contract?.Input || "").toLowerCase();
    const BillableCTC = associate?.BillableCTC;

    const vars = extractVariables(c);

    const wrapped = `
      ${vars.map(v => `var ${v};`).join("\n")}

      ${c}

      const __result = {};
      ${vars.map(v => `if (typeof ${v} !== "undefined") __result["${v}"] = ${v};`).join("\n")}

      __result;
    `;

    return eval(wrapped);

  } catch (e) {
    return `Error: ${e.message}`;
  }
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
 
//     // 🔥 IMPORTANT FIX
//    const Input = (contract?.Input || "").toLowerCase();
//    console.log(Input)
//     const BillableCTC = associate?.BillableCTC;
 
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
  var MarkUp, ReplacementPeriod, ExecutedServiceTerms;

  ${c}

  const __result = {};
  if (typeof MarkUp !== "undefined") __result["MarkUp"] = MarkUp;
  if (typeof ReplacementPeriod !== "undefined") __result["ReplacementPeriod"] = ReplacementPeriod;
  if (typeof ExecutedServiceTerms !== "undefined") __result["ExecutedServiceTerms"] = ExecutedServiceTerms;

  __result;
`;

    return eval(wrapped);

  } catch (e) {
    return `Error: ${e.message}`;
  }
}

 function extractInputs(script) {
  // Clean the input string (assuming you have a clean function to sanitize or pre-process it)
  const cleaned = clean(script);
  // Define the keywords as a regex pattern
  const keywordsPattern = [
    'mpg', 'tc', 'olda', 'newa', 'contractualstaffing', 'cr', 'i', 'ni', 'blu', 'wht', 'bea', 'bec', 'casm',
    'ipod', 'fp', 'o', 'p1', 'p2', '0-4y','>4.1y','e&b', 'ae', 'admin', 'contractualhiring', 'permanenthiring', 'gsc', 'nongsc',
    'r&df', 'referral', 'grade1', 'grade2', 'ar/ds', 'grade3', 'grade4', 'grade5', 'abovegrade6', 'grade1f', 'grade2f',
    'grade3f', 'grade4f', 'grade5f', 'abovegrade6f', 'headhuntingposition', 'non-headhuntingposition', 'belowgrade7&piisc', 
    '0-1y', '1.1-2y', '2.1-3y', '3.1-4y', '4.1-5y', '5.1-6y', '6.1-7y', '7.1-8y', '8.1-9y', '9.1-10y', '10.1-11y', 'rpo_partner_band1_uc',
    '11.1-12y', '12.1-13y', '13.1-14y', '14.1-15y', '15.1-16y', '16.1-17y', '17.1-18y', '18.1-19y', '19.1-20y', '>20y','bandaav','bandaanv','bandb2' ,'bandb3','bandc1','bandc2','bandd1above','rpo_partner_band2','rpo_partner_band3_and_band4','rpo_partner_band5_and_band6','sbic_hired_branch_band1_uc','sbic_hired_branch_band2',
    'bandh', 'bandi', 'bandj', 'bandk','band3.1', 'band3.2', 'band4.1', 'band4.2', 'band5.1', 'band5.2', 'band5.3', 'band5.4', 'band6.1', 'c1-c2','c3-c5','c6-c8',
    'band6.2', 'band6.3', 'band6.4', 'band6.5', 'band5.2f', 'band5.3f', 'band5.4f', 'band6.1f', 'band6.2f', 'band6.3f', 'sbic_hired_branch_band3_to_band6','fresher_or_intern_or_apprentice',
    'band6.4f', 'band6.5f', 'bandl', 'band1', 'band2', 'band3', 'band4', 'band5', 'band6', 'band7', 'band8', 
    'band9', 'band10', 'band11', 'banda', 'bandb', 'bandb1', 'bandc', 'bandd1', 'bandh', 'bande', 'band28', 'band40', 'band30',
    'belowgrade7&referral', 'belowgrade7&rehire', 'belowgrade7&rpo', 'belowgrade7&mpg', 'abovegrade8&piisc', 'abovegrade8&referral',
    'abovegrade8&rehire','lc_strategicband','lc_executiveband', 'abovegrade8&rpo', 'abovegrade8&mpg', 'juniormanagement', 'middlemanagement', 'seniormanagement',
    'referralrpo', 'uptodirector', 'avp', 'fresher', 'contractor', 'retainer', 'f', 'rpo', 'perm', 'bandl4', 'bandl5', 'bandl6', 
    'bandl7', 'bandl8',  'sa/sco', 'pe', 'tl', 'avp', 'vp', 'bandl9', 'bandl10', 'bandl11', 'bandl12', 'bandl13', 'cops', 
    'o', 'woretainer','mmc_professionalband','jmc_operationalband','jmc_administrativeband',, 'va', 'b1', 'b2', 'c1', 'c2', 'd1', 'd2', 'n grade a', 'nlr', 'lr', 'retainer', 'bandpabove&sc/st', 
    'bandpabove&t','smc_tacticalband', 'bandpabove&f', 'n b1', 'n b2', 'n c1', 'n c2', 'n d1', 'n d2', 'bandl4', 'bandl5', 'bandl6', 'bandl7',
    'bandl8', 'bandl9', 'bandl10', 'bandl11', 'bandl12', 'bandl13', 'bandpabove&h', 'bandrq&sc/st', 'bandrq&t', 'bandrq&f', 
    'bandrq&h', '1-5hires','dl','idl_lessthan_am','referral','diversity_dl_dlt','dl_t','idl_m','dl_lessthan_am','idl _am_dm','6-10hires','16&ahires','11-15hires','bandts&sc/st', 'bandts&t', 'bandts&f', 'bandts&h', 'am', 'manager', 'ad', 'principal', 'director', 'partner','neg1','neg2','neg3','neg4','neg5','neg6','neg7','ee','eese','eem1','eem2','eem3','vpm4','vpm5','vpg1','vpg2'
  ];

  // Combine the keywords into a single regular expression with word boundaries and case insensitivity
  const regex = new RegExp(`\\b(${keywordsPattern.join('|')})\\b`, 'g');  // 'g' for global matching

  // Match all occurrences of the keywords
  const matches = cleaned.match(regex);

  // Return unique matches (if any)
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
    { name: "perm", script: get(row, COL.payroll), terms: get(row, COL.payrollTerms),replacementPeriod: get(row, COL.replacementPeriod) }
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
      // result = {
      //   output: run(r.script, baseData)
      // };
      result = run(r.script, baseData);
    }
 
    results.push({
      service: r.name,
      terms: r.terms,
      script: r.script,
      replacementPeriod: r.replacementPeriod,
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
// fs.writeFileSync("goliveoutput3.json", JSON.stringify(output, null, 2));
// fs.writeFileSync("output3.json", JSON.stringify(output, null, 2));
fs.writeFileSync("output1.json", JSON.stringify(output, null, 2));
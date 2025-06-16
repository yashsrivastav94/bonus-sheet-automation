function automateNextBonusSheetImproved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const config = {
    sourceSheetName: "", // auto-detected
    targetSheetName: "", // auto-determined
    dataCollectorSheetName: "Data collector",
    contractorSpreadsheetId: "<YOUR_CONTRACTOR_SPREADSHEET_ID>", // TODO: Replace with your Sheet ID
    contractorSheetName: "Contractors",
    dropdownColumns: [2, 7],
    dataCollectorBlocks: [
      [1, 3], [5, 3], [11, 3], [16, 3]
    ],
    dataCollectorOffsets: {
      bonusType: 1,
      namePartsStart: 8,
      namePartsNumCols: 2,
      interviewerName: 14
    },
    formulaSourceRange: { row: 2, col: 5, numRows: 1, numCols: 4 },
    dryRun: false
  };

  function parseSheetMonthYear(name) {
    const [month, year] = name.split(" ");
    const monthIndex = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"].indexOf(month);
    return monthIndex !== -1 && !isNaN(parseInt(year)) ? new Date(year, monthIndex, 1) : null;
  }

  let latestDate = null, latestSheetName = null;
  ss.getSheets().forEach(s => {
    const d = parseSheetMonthYear(s.getName());
    if (d && (!latestDate || d > latestDate)) {
      latestDate = d;
      latestSheetName = s.getName();
    }
  });
  if (!latestSheetName) return ui.alert("No valid 'Month YYYY' sheet found.");

  const nextDate = new Date(latestDate.getFullYear(), latestDate.getMonth() + 1, 1);
  const nextMonthLongName = nextDate.toLocaleString("en-US", { month: "long", year: "numeric" });
  const nextMonthShort = nextDate.toLocaleString("en-US", { month: "short" });
  const nextYear = nextDate.getFullYear().toString();

  config.sourceSheetName = latestSheetName;
  config.targetSheetName = nextMonthLongName;

  if (config.dryRun) return ui.alert(`Dry Run: Next = "${config.targetSheetName}" from "${config.sourceSheetName}"`);

  if (ss.getSheetByName(config.targetSheetName)) return ui.alert(`Sheet "${config.targetSheetName}" already exists.`);

  const source = ss.getSheetByName(config.sourceSheetName);
  const collector = ss.getSheetByName(config.dataCollectorSheetName);
  if (!source || !collector) return;

  for (const ref of ["C1", "G1", "M1", "R1"]) {
    const val = collector.getRange(ref).getDisplayValue().toLowerCase();
    const clean = val.replace(/\s+/g, " ");
    const valid = (
      clean.includes(nextMonthLongName.toLowerCase().split(" ")[0]) && (clean.includes(nextYear) || clean.includes(nextYear.slice(-2))) ||
      clean.includes(nextMonthShort.toLowerCase()) && (clean.includes(nextYear) || clean.includes(nextYear.slice(-2)))
    );
    if (!valid) return ui.alert(`Mismatch in ${ref}: "${val}" ≠ ${nextMonthLongName}`);
  }

  const filter = source.getFilter();
  if (filter) filter.remove();

  const idx = ss.getSheets().findIndex(s => s.getName() === config.sourceSheetName);
  const newSheet = ss.insertSheet(config.targetSheetName, idx);
  source.getRange(1, 1, 1, source.getMaxColumns()).copyTo(newSheet.getRange(1, 1));
  const row2Format = source.getRange(2, 1, 1, source.getMaxColumns());

  config.dropdownColumns.forEach(col => {
    const rule = source.getRange(2, col).getDataValidation();
    if (rule) newSheet.getRange(2, col, newSheet.getMaxRows() - 1).setDataValidation(rule);
  });

  let targetRow = 2, stacked = [], referralNotes = {}, interviewerNotes = {}, signinNotes = {};
  config.dataCollectorBlocks.forEach(([startCol, numCols]) => {
    const max = collector.getLastRow() - 2;
    if (max <= 0) return;

    const bonusTypeCol = startCol + config.dataCollectorOffsets.bonusType;
    const namePartsCol = config.dataCollectorOffsets.namePartsStart;
    const interviewerCol = config.dataCollectorOffsets.interviewerName;

    const data = collector.getRange(3, startCol, max, numCols).getValues();
    const typeData = collector.getRange(3, bonusTypeCol, max, 1).getValues();
    const nameParts = collector.getRange(3, namePartsCol, max, config.dataCollectorOffsets.namePartsNumCols).getValues();
    const interviewee = collector.getRange(3, interviewerCol, max, 1).getValues();

    data.forEach((row, i) => {
      if (row.some(c => c !== "")) {
        stacked.push(row);
        const type = typeData[i][0];
        const rowNum = stacked.length - 1 + targetRow;
        if (type === "Referral bonus (AE)") referralNotes[rowNum] = { name: nameParts[i][1], part: nameParts[i][0] };
        if (type === "Interviewer bonus") interviewerNotes[rowNum] = interviewee[i][0];
        if (type === "Sign-in bonus") signinNotes[rowNum] = nameParts[i][0];
      }
    });
  });

  if (stacked.length) {
    newSheet.getRange(targetRow, 1, stacked.length, stacked[0].length).setValues(stacked);
    targetRow += stacked.length;
  }

  const contractorSS = SpreadsheetApp.openById(config.contractorSpreadsheetId);
  const contractorSheet = contractorSS.getSheetByName(config.contractorSheetName);
  const contractorData = contractorSheet.getRange("A3:K").getValues().filter(r => r[0]);
  const contractorMap = Object.fromEntries(contractorData.map(r => [r[0].toString().trim(), r[10]]));
  const names = contractorData.map(r => [r[0]]);
  newSheet.getRange(targetRow, 1, names.length, 1).setValues(names);
  targetRow += names.length;

  const last = targetRow - 1;
  if (last >= 2) row2Format.copyTo(newSheet.getRange(2, 1, last - 1, source.getMaxColumns()), { formatOnly: true });

  const nameData = newSheet.getRange(2, 1, newSheet.getLastRow() - 1, 1).getValues();
  const formulaRow = source.getRange(config.formulaSourceRange.row, config.formulaSourceRange.col, 1, config.formulaSourceRange.numCols);
  const formulas = formulaRow.getFormulasR1C1()[0];
  const toApply = nameData.map(r => r[0] ? formulas : Array(formulas.length).fill(""));
  newSheet.getRange(2, config.formulaSourceRange.col, toApply.length, config.formulaSourceRange.numCols).setFormulasR1C1(toApply);

  const bValues = newSheet.getRange(2, 2, nameData.length, 1).getValues();
  const dValues = [];

  for (let i = 0; i < bValues.length; i++) {
    const bVal = String(bValues[i][0]).trim();
    const name = String(nameData[i][0]).trim();
    let dVal = "";

    if (bVal === "MRKT bonus monthly") dVal = `${nextMonthLongName.split(" ")[0]} Marketing Bonus`;
    else if (bVal === "RCRT bonus monthly") dVal = `${nextMonthLongName.split(" ")[0]} Recruiting Bonus`;
    else if (bVal === "HR bonus quarterly") dVal = `${nextMonthLongName.split(" ")[0]} HR Bonus`;
    else if (bVal === "ENG bonus quarterly") dVal = `${nextMonthLongName.split(" ")[0]} Engineering Bonus`;
    else if (["Self-Development/Education", "Self-Development/Team Sports", "Self-Development/Individual Sports", "Self-Development/Co-working"].includes(bVal)) {
      dVal = bVal;
    } else if (bVal === "Teambuilding") {
      const dept = contractorMap[name] || "Department";
      dVal = `${dept} Teambuilding`;
    } else if (bVal === "Referral bonus (AE)" && referralNotes[i + 2]) {
      const r = referralNotes[i + 2];
      dVal = `Referral bonus (AE) for ${r.name} (AE) ${r.part} / 2`;
    } else if (bVal === "Interviewer bonus" && interviewerNotes[i + 2]) {
      dVal = `Bonus for Hiring ${interviewerNotes[i + 2]}`;
    } else if (bVal === "Sign-in bonus" && signinNotes[i + 2]) {
      dVal = `Sign-in bonus ${signinNotes[i + 2]} / 2`;
    }

    dValues.push([dVal]);
  }

  if (dValues.length) newSheet.getRange(2, 4, dValues.length, 1).setValues(dValues);
  SpreadsheetApp.flush();
  ui.alert(`Bonus Sheet "${config.targetSheetName}" created successfully!`);
}

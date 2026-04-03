function snapshotNetWorth() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Net-Worth");
  const target = ss.getSheetByName("Net-Worth-Evolution");

  /* ---------- CREATE HEADERS IF MISSING ---------- */

  const headers = [
    "Date",
    "T-Bills",
    "Real Estate",
    "IBKR",
    "Deposit MAIB USD",
    "Revolut EUR",
    "Revolut USD",
    "Revolut RON",
    "Cash MAIB MDL",
    "Cash MAIB EUR",
    "Cash VB MDL",
    "Car Huyndai Tucson 2019",
    "Liquid",
    "Illiquid",
    "Total",
    "Change",
    "Percent"
  ];

  if (target.getLastRow() === 0) {

    target.getRange(1,1,1,headers.length).setValues([headers]);

    target.getRange(1,1,1,headers.length)
      .setFontWeight("bold")
      .setBackground("#f1f3f4");

    target.getRange(1,1,2,headers.length).applyRowBanding();

    }

    /* ---------- LOAD DATA ---------- */

    const rows = source.getRange(2,1,source.getLastRow()-1,3).getValues();

    const assets = {
      "T-Bills":0,
      "Real Estate":0,
      "IBKR":0,
      "Deposit MAIB USD":0,
      "Revolut EUR":0,
      "Revolut USD":0,
      "Revolut RON":0,
      "Cash MAIB MDL":0,
      "Cash MAIB EUR":0,
      "Cash VB MDL":0,
      "Car Huyndai Tucson 2019":0
    };

    rows.forEach(r => {

    const name = String(r[0]).toLowerCase();
    const value = Number(r[2]) || 0;

    if(name.includes("t-bill")) assets["T-Bills"] += value;

    else if(
      name.includes("real estate") ||
      name.includes("apartment") ||
      name.includes("house") ||
      name.includes("property") ||
      name.includes("land") ||
      name.includes("parking")
    ){
      assets["Real Estate"] += value;
    }

    else if(name.includes("ibkr")) assets["IBKR"] += value;

    else if(name.includes("deposit")) assets["Deposit MAIB USD"] += value;

    else if(name.includes("revolut eur")) assets["Revolut EUR"] += value;

    else if(name.includes("revolut usd")) assets["Revolut USD"] += value;

    else if(name.includes("revolut ron")) assets["Revolut RON"] += value;

    else if(name.includes("cash maib mdl")) assets["Cash MAIB MDL"] += value;

    else if(name.includes("cash maib eur")) assets["Cash MAIB EUR"] += value;

    else if(name.includes("cash vb mdl")) assets["Cash VB MDL"] += value;

    else if(name.includes("car")) assets["Car Huyndai Tucson 2019"] += value;

  });

  /* ---------- CALCULATIONS ---------- */

  const liquid =
  assets["T-Bills"] +
  assets["IBKR"] +
  assets["Deposit MAIB USD"] +
  assets["Revolut EUR"] +
  assets["Revolut USD"] +
  assets["Revolut RON"] +
  assets["Cash MAIB MDL"] +
  assets["Cash MAIB EUR"] +
  assets["Cash VB MDL"];

  const illiquid =
  assets["Real Estate"] +
  assets["Car Huyndai Tucson 2019"];

  const total = liquid + illiquid;

  const today = new Date();
  const firstDay = new Date(today.getFullYear(),today.getMonth(),1);

  const lastRow = target.getLastRow();
  
  let writeRow = lastRow + 1;

  if (lastRow > 1) {

    const lastDate = target.getRange(lastRow,1).getValue();

    if (
      lastDate instanceof Date &&
      lastDate.getMonth() === firstDay.getMonth() &&
      lastDate.getFullYear() === firstDay.getFullYear()
    ) {
      writeRow = lastRow; // overwrite instead of adding
    }
  }

  /* ---------- CHANGE CALCULATION ---------- */
  let prevTotal = null;

  const prevRow = writeRow - 1;

  if (prevRow >= 2) {
    prevTotal = target.getRange(prevRow,15).getValue();
  }

  const change = prevTotal !== null ? total - prevTotal : 0;
  const percent = prevTotal ? change / prevTotal : 0;

  /* ---------- ADD ROW ---------- */

  target.getRange(writeRow,1,1,17).setValues([[
    firstDay,
    assets["T-Bills"],
    assets["Real Estate"],
    assets["IBKR"],
    assets["Deposit MAIB USD"],
    assets["Revolut EUR"],
    assets["Revolut USD"],
    assets["Revolut RON"],
    assets["Cash MAIB MDL"],
    assets["Cash MAIB EUR"],
    assets["Cash VB MDL"],
    assets["Car Huyndai Tucson 2019"],
    liquid,
    illiquid,
    total,
    change,
    percent
  ]]);

  target.getRange(writeRow, 17).setNumberFormat("0.00%");

  formatEvolutionTable(target);
}

function formatEvolutionTable(sheet) {

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const totalRange   = sheet.getRange(2, 15, lastRow - 1, 1); // O
  const changeRange  = sheet.getRange(2, 16, lastRow - 1, 1); // P
  const percentRange = sheet.getRange(2, 17, lastRow - 1, 1); // Q

  const existingRules = sheet.getConditionalFormatRules();

  const filteredRules = existingRules.filter(rule => {
    return !rule.getRanges().some(r => {
      const col = r.getColumn();
      return col === 15 || col === 16 || col === 17;
    });
  });

  /* ---------- TOTAL (compare with previous row) ---------- */

  const totalPositive =
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$O2>$O1')
      .setBackground("#d4edda")
      .setFontColor("#1e7e34")
      .setRanges([totalRange])
      .build();

  const totalNegative =
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$O2<$O1')
      .setBackground("#f8d7da")
      .setFontColor("#c82333")
      .setRanges([totalRange])
      .build();

  /* ---------- CHANGE + PERCENT ---------- */

  const positiveRule =
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#d4edda")
      .setFontColor("#1e7e34")
      .setRanges([changeRange, percentRange])
      .build();

  const negativeRule =
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground("#f8d7da")
      .setFontColor("#c82333")
      .setRanges([changeRange, percentRange])
      .build();

  sheet.setConditionalFormatRules([
    ...filteredRules,
    totalPositive,
    totalNegative,
    positiveRule,
    negativeRule
  ]);
}

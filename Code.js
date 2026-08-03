function snapshotNetWorth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName("Net-Worth");
  const target = ss.getSheetByName("Net-Worth-Evolution");

  ensureHeaders(target);

  const rows = getSourceRows(source);
  const assets = computeAssets(rows);

  const { liquid, illiquid } = computeTotals(assets);
  const totals = getSourceTotals(source);

  const firstDay = getFirstDayOfMonth();
  const writeRow = getWriteRow(target, firstDay);

  const prev = getPreviousTotals(target, writeRow);

  const eurChange = computeChanges(totals.eur, prev.eur);
  const mdlChange = computeChanges(totals.mdl, prev.mdl);
  const usdChange = computeChanges(totals.usd, prev.usd);

  writeSnapshot(
  target,
  writeRow,
  firstDay,
  assets,
  liquid,
  illiquid,
  totals,
  eurChange,
  mdlChange,
  usdChange
  );

  colorizeRow(target, writeRow);
}

function getSourceTotals(source) {
  const lastRow = source.getLastRow();

  const mdl = source.getRange(lastRow, 2).getValue();
  const eur = source.getRange(lastRow, 3).getValue();
  const usd = source.getRange(lastRow, 4).getValue();

  return { mdl, eur, usd };
}

function getPreviousTotals(sheet, writeRow) {
  const prevRow = writeRow - 1;
  if (prevRow < 2) return { eur: null, mdl: null, usd: null };

  return {
    eur: sheet.getRange(prevRow, 15).getValue(),
    mdl: sheet.getRange(prevRow, 16).getValue(),
    usd: sheet.getRange(prevRow, 17).getValue()
  };
}

function computeChanges(current, previous) {
  const change = previous !== null ? current - previous : 0;
  const percent = previous ? change / previous : 0;
  return { change, percent };
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() !== 0) return;

  const headers = [
    "Date","T-Bills","Real Estate","IBKR","Deposit MAIB USD",
    "Revolut EUR","Revolut USD","Revolut RON",
    "Cash MAIB MDL","Cash MAIB EUR","Cash VB MDL",
    "Car Huyndai Tucson 2019",
    "Liquid","Illiquid",
    "Total EUR","Total MDL","Total USD",
    "Change EUR","Percent EUR",
    "Change MDL","Percent MDL",
    "Change USD","Percent USD"
  ];

  sheet.getRange(1,1,1,headers.length).setValues([headers]);
  sheet.getRange(1,1,1,headers.length)
    .setFontWeight("bold")
    .setBackground("#f1f3f4");

  sheet.getRange(1,1,2,headers.length).applyRowBanding();
}

function getSourceRows(source) {
  return source
    .getRange(2,1,source.getLastRow()-1,3)
    .getValues();
}

function computeAssets(rows) {
  const assets = {
    "T-Bills":0,"Real Estate":0,"IBKR":0,"Deposit MAIB USD":0,
    "Revolut EUR":0,"Revolut USD":0,"Revolut RON":0,
    "Cash MAIB MDL":0,"Cash MAIB EUR":0,"Cash VB MDL":0,
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
    ) assets["Real Estate"] += value;

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

  return assets;
}

function computeTotals(assets) {
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

  return {
    liquid,
    illiquid
  };
}

function getFirstDayOfMonth() {
  const today = new Date();
  return new Date(today.getFullYear(), today.getMonth(), 1);
}

function getWriteRow(sheet, firstDay) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return lastRow + 1;

  const lastDate = sheet.getRange(lastRow,1).getValue();

  if (
    lastDate instanceof Date &&
    lastDate.getMonth() === firstDay.getMonth() &&
    lastDate.getFullYear() === firstDay.getFullYear()
  ) {
    return lastRow; // overwrite
  }

  return lastRow + 1;
}

function writeSnapshot(
  sheet,
  row,
  date,
  assets,
  liquid,
  illiquid,
  totals,
  eurChange,
  mdlChange,
  usdChange
) {
  sheet.getRange(row, 1, 1, 23).setValues([[
    date,
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
    totals.eur, // 15
    totals.mdl, // 16
    totals.usd, // 17
    eurChange.change,   // 18
    eurChange.percent,  // 19
    mdlChange.change,   // 20
    mdlChange.percent,  // 21
    usdChange.change,   // 22
    usdChange.percent   // 23
  ]]);

  // format percentages
  sheet.getRange(row, 19).setNumberFormat("0.00%");
  sheet.getRange(row, 21).setNumberFormat("0.00%");
  sheet.getRange(row, 23).setNumberFormat("0.00%");
}

function colorizeRow(sheet, row) {
  if (row <= 2) return; // no previous row to compare

  const numCols = 23;
  const currentValues = sheet.getRange(row, 1, 1, numCols).getValues()[0];
  const prevValues = sheet.getRange(row - 1, 1, 1, numCols).getValues()[0];

  for (let col = 2; col <= numCols; col++) {
    const current = Number(currentValues[col - 1]);
    const prev = Number(prevValues[col - 1]);

    if (isNaN(current)) continue;

    const cell = sheet.getRange(row, col);

    // reset formatting
    cell.setBackground(null).setFontColor(null);

    // Change / Percent columns: compare against 0
    if ([18, 19, 20, 21, 22, 23].includes(col)) {
      if (current > 0) {
        cell.setBackground("#d4edda").setFontColor("#1e7e34");
      } else if (current < 0) {
        cell.setBackground("#f8d7da").setFontColor("#c82333");
      }
      continue;
    }

    // All other columns: compare against previous month
    if (isNaN(prev)) continue;

    if (current > prev) {
      cell.setBackground("#d4edda").setFontColor("#1e7e34");
    } else if (current < prev) {
      cell.setBackground("#f8d7da").setFontColor("#c82333");
    }
  }
}

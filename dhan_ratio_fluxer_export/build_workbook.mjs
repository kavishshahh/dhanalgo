import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "C:/Users/kavis/OneDrive/Desktop";
const inputPath = `${root}/ratio_fluxer_trades_raw.json`;
const outputPath = `${root}/Ratio-Fluxer_Past_Trades_2025-09-18_to_2026-09-18.xlsx`;
const previewDir = `${root}/dhan_ratio_fluxer_export/previews`;
const sourceUrl = "https://algos.dhan.co/managers/stratzy/ratio-fluxer-credit-spread-expiry/698075eaf867bf12b20442e5?tab=past-trades";
const strategyName = "Ratio-Fluxer Credit Spread Expiry";
const fromDate = "2025-09-18";
const toDate = "2026-09-18";

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
if (payload.status !== "success" || !Array.isArray(payload.data)) {
  throw new Error(`Unexpected API response: ${payload.message || payload.status}`);
}
const trades = payload.data.slice().sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));

const font = "Arial";
const colors = {
  title: "#261A5A",
  header: "#3F2B96",
  header2: "#5B46B2",
  light: "#EEEAFB",
  lighter: "#F7F5FD",
  text: "#1F2937",
  muted: "#667085",
  line: "#D9D5E8",
  green: "#087A55",
  greenFill: "#E6F4EE",
  red: "#B42318",
  redFill: "#FDECEC",
  amber: "#B54708",
};

function istDate(iso) {
  if (!iso) return null;
  const source = new Date(iso);
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric", month: "2-digit", day: "2-digit",
    hour: "2-digit", minute: "2-digit", second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(source).reduce((acc, p) => (acc[p.type] = p.value, acc), {});
  return new Date(Number(parts.year), Number(parts.month) - 1, Number(parts.day), Number(parts.hour), Number(parts.minute), Number(parts.second));
}

function monthStartFromIso(iso) {
  const d = istDate(iso);
  return new Date(d.getFullYear(), d.getMonth(), 1);
}

function monthKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function legPnlPerLot(leg) {
  const direction = leg.transactionType === "SELL" ? 1 : -1;
  return direction * (Number(leg.entryPrice || 0) - Number(leg.exitPrice || 0)) * Number(leg.quantity || 0) * Number(leg.multiplier || 0);
}

function round2(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}

function excelCol(n) {
  let s = "";
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}

function styleBase(sheet, usedRange) {
  sheet.showGridLines = false;
  usedRange.format.font = { name: font, size: 10, color: colors.text };
  usedRange.format.verticalAlignment = "center";
}

function styleTitle(sheet, title, subtitle, lastCol) {
  const end = excelCol(lastCol);
  sheet.getRange(`A2:${end}2`).format.borders = { bottom: { style: "thin", color: colors.header } };
  sheet.getRange("A2").values = [[title]];
  sheet.getRange("A2").format.font = { name: font, size: 15, bold: true, color: colors.title };
  sheet.getRange("A3").values = [[subtitle]];
  sheet.getRange("A3").format.font = { name: font, size: 9, italic: true, color: colors.muted };
}

function styleHeader(range, fill = colors.header) {
  range.format = {
    fill,
    font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "inside", style: "thin", color: "#FFFFFF" },
  };
  range.format.rowHeight = 32;
}

function addOutcomeFormatting(range, firstDataRow, outcomeColLetter) {
  range.conditionalFormats.addCustom(`=$${outcomeColLetter}${firstDataRow}="Win"`, {
    fill: colors.greenFill, font: { color: colors.green, bold: true },
  });
  range.conditionalFormats.addCustom(`=$${outcomeColLetter}${firstDataRow}="Loss"`, {
    fill: colors.redFill, font: { color: colors.red, bold: true },
  });
}

const workbook = Workbook.create();
const summarySheet = workbook.worksheets.add("Monthly Summary");
const tradesSheet = workbook.worksheets.add("Trades");
const legsSheet = workbook.worksheets.add("Trade Legs");
const exitsSheet = workbook.worksheets.add("Exit Details");
summarySheet.tabColor = colors.title;
tradesSheet.tabColor = colors.header2;
legsSheet.tabColor = colors.header2;
exitsSheet.tabColor = colors.header2;

// Monthly Summary
styleTitle(summarySheet, `${strategyName} - past trades`, `One-year window: ${fromDate} to ${toDate}. Times are India Standard Time (IST).`, 17);
summarySheet.getRange("A4").values = [[`Source: ${sourceUrl}`]];
summarySheet.getRange("A4").format.font = { name: font, size: 9, color: colors.muted };
summarySheet.getRange("A5").values = [["Dhan lists the strategy as live since 09-Mar-2026; its historical endpoint also returns earlier records beginning 22-Sep-2025."]];
summarySheet.getRange("A5").format.font = { name: font, size: 9, italic: true, color: colors.amber };

const monthStarts = [];
for (let d = new Date(2025, 8, 1); d <= new Date(2026, 8, 1); d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) monthStarts.push(new Date(d));
const monthlyRows = monthStarts.map(month => {
  const key = monthKey(month);
  const entered = trades.filter(t => monthKey(monthStartFromIso(t.createdAt)) === key);
  const exited = trades.filter(t => monthKey(monthStartFromIso(t.closedOn)) === key);
  const pnlEntry = entered.reduce((s, t) => s + Number(t.pnl || 0), 0);
  const pnlExit = exited.reduce((s, t) => s + Number(t.pnl || 0), 0);
  const wins = entered.filter(t => Number(t.pnl || 0) > 0).length;
  const losses = entered.filter(t => Number(t.pnl || 0) < 0).length;
  const avg = entered.length ? pnlEntry / entered.length : 0;
  const best = entered.length ? Math.max(...entered.map(t => Number(t.pnl || 0))) : 0;
  const worst = entered.length ? Math.min(...entered.map(t => Number(t.pnl || 0))) : 0;
  return [month.toLocaleString("en-US", { month: "short", year: "2-digit" }).replace(" ", "-"), entered.length, exited.length, round2(pnlEntry), round2(pnlExit), wins, losses, entered.length ? wins / entered.length : null, round2(avg), round2(best), round2(worst)];
});
const monthlyHeaders = ["Month", "Trades entered", "Trades exited", "P&L by entry month (INR)", "P&L by exit month (INR)", "Wins", "Losses", "Win rate", "Average P&L (INR)", "Best trade (INR)", "Worst trade (INR)"];
summarySheet.getRange("A7:K7").values = [monthlyHeaders];
summarySheet.getRange(`A8:K${7 + monthlyRows.length}`).values = monthlyRows;
styleHeader(summarySheet.getRange("A7:K7"));
summarySheet.tables.add(`A7:K${7 + monthlyRows.length}`, true, "MonthlySummaryTable").style = "TableStyleMedium4";
styleHeader(summarySheet.getRange("A7:K7"));
summarySheet.getRange(`A8:A${7 + monthlyRows.length}`).format.numberFormat = "mmm-yy";
summarySheet.getRange(`D8:E${7 + monthlyRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
summarySheet.getRange(`H8:H${7 + monthlyRows.length}`).format.numberFormat = "0.0%";
summarySheet.getRange(`I8:K${7 + monthlyRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";

const totalPnl = round2(trades.reduce((s, t) => s + Number(t.pnl || 0), 0));
const wins = trades.filter(t => Number(t.pnl || 0) > 0).length;
const losses = trades.filter(t => Number(t.pnl || 0) < 0).length;
const firstTrade = trades.length ? istDate(trades[0].createdAt) : null;
const lastTrade = trades.length ? istDate(trades[trades.length - 1].createdAt) : null;
summarySheet.getRange("M2:N2").values = [["Metric", "Value"]];
summarySheet.getRange("M3:N9").values = [
  ["Trades", trades.length],
  ["Option legs", trades.reduce((s, t) => s + (t.trades?.length || 0), 0)],
  ["Total P&L", totalPnl],
  ["Wins", wins],
  ["Losses", losses],
  ["Win rate", trades.length ? wins / trades.length : null],
  ["Entry-date coverage", `${firstTrade ? firstTrade.toLocaleDateString("en-IN") : "n.a."} to ${lastTrade ? lastTrade.toLocaleDateString("en-IN") : "n.a."}`],
];
styleHeader(summarySheet.getRange("M2:N2"), colors.header2);
summarySheet.getRange("M3:M9").format.font = { name: font, size: 10, bold: true, color: colors.text };
summarySheet.getRange("N5").format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
summarySheet.getRange("N8").format.numberFormat = "0.0%";
summarySheet.getRange("M2:N9").format.borders = { preset: "outside", style: "thin", color: colors.line };
summarySheet.getRange("M3:N9").format.fill = colors.lighter;

const chart = summarySheet.charts.add("bar", [summarySheet.getRange(`A7:A${7 + monthlyRows.length}`), summarySheet.getRange(`D7:D${7 + monthlyRows.length}`)]);
chart.title = "Monthly P&L by entry month";
chart.titleTextStyle.typeface = font;
chart.titleTextStyle.fontSize = 12;
chart.hasLegend = false;
chart.xAxis = { axisType: "textAxis", textStyle: { typeface: font, fontSize: 9 } };
chart.yAxis = { numberFormatCode: "₹#,##0", numberFormatSourceLinked: false, textStyle: { typeface: font, fontSize: 9 } };
chart.setPosition("M11", "T27");
summarySheet.freezePanes.freezeRows(7);
summarySheet.getRange("A:K").format.columnWidth = 16;
summarySheet.getRange("A:A").format.columnWidth = 12;
summarySheet.getRange("D:E").format.columnWidth = 20;
summarySheet.getRange("I:K").format.columnWidth = 18;
summarySheet.getRange("M:M").format.columnWidth = 22;
summarySheet.getRange("N:N").format.columnWidth = 23;
styleBase(summarySheet, summarySheet.getUsedRange());

// Trades
const tradeHeaders = ["Trade no.", "Entry month", "Exit month", "Trade name", "Signal ID", "Entry time (IST)", "Exit time (IST)", "Duration (hours)", "Total P&L (INR)", "Total P&L %", "Per-lot P&L (INR)", "Per-lot P&L %", "Margin required (INR)", "Capital allocated %", "Lots bought", "Option legs", "Exit summaries", "Exit events", "Calculated leg P&L (INR)", "P&L difference (INR)", "Outcome"];
const tradeRows = trades.map((t, idx) => {
  const entry = istDate(t.createdAt), exit = istDate(t.closedOn);
  const calcLegPnl = round2((t.trades || []).reduce((s, leg) => s + legPnlPerLot(leg) * Number(t.lotsBought || 0), 0));
  return [
    idx + 1, monthStartFromIso(t.createdAt), monthStartFromIso(t.closedOn), t.name, t.signalId,
    entry, exit, (exit - entry) / 36e5, Number(t.pnl || 0), Number(t.pnlPercent || 0) / 100,
    Number(t.signalPnl || 0), Number(t.signalPnlPercent || 0) / 100, Number(t.marginRequired || 0), Number(t.capitalAllocated || 0) / 100,
    Number(t.lotsBought || 0), t.trades?.length || 0, t.summary?.length || 0, t.exits?.length || 0,
    calcLegPnl, round2(Number(t.pnl || 0) - calcLegPnl), Number(t.pnl || 0) > 0 ? "Win" : Number(t.pnl || 0) < 0 ? "Loss" : "Flat",
  ];
});
styleTitle(tradesSheet, "Strategy trades", `Source: ${sourceUrl}`, tradeHeaders.length);
tradesSheet.getRange(`A5:${excelCol(tradeHeaders.length)}5`).values = [tradeHeaders];
tradesSheet.getRange(`A6:${excelCol(tradeHeaders.length)}${5 + tradeRows.length}`).values = tradeRows;
styleHeader(tradesSheet.getRange(`A5:${excelCol(tradeHeaders.length)}5`));
tradesSheet.tables.add(`A5:${excelCol(tradeHeaders.length)}${5 + tradeRows.length}`, true, "TradesTable").style = "TableStyleMedium4";
styleHeader(tradesSheet.getRange(`A5:${excelCol(tradeHeaders.length)}5`));
tradesSheet.getRange(`B6:C${5 + tradeRows.length}`).format.numberFormat = "mmm-yy";
tradesSheet.getRange(`F6:G${5 + tradeRows.length}`).format.numberFormat = "dd-mmm-yyyy hh:mm";
tradesSheet.getRange(`H6:H${5 + tradeRows.length}`).format.numberFormat = "0.0";
for (const col of ["I", "K", "M", "S", "T"]) tradesSheet.getRange(`${col}6:${col}${5 + tradeRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
for (const col of ["J", "L", "N"]) tradesSheet.getRange(`${col}6:${col}${5 + tradeRows.length}`).format.numberFormat = "0.00%";
addOutcomeFormatting(tradesSheet.getRange(`A6:${excelCol(tradeHeaders.length)}${5 + tradeRows.length}`), 6, "U");
tradesSheet.freezePanes.freezeRows(5);
tradesSheet.freezePanes.freezeColumns(1);
tradesSheet.getUsedRange().format.autofitColumns();
tradesSheet.getRange("D:D").format.columnWidth = 42;
tradesSheet.getRange("E:E").format.columnWidth = 25;
tradesSheet.getRange("F:G").format.columnWidth = 21;
tradesSheet.getRange("A:U").format.rowHeight = 22;
tradesSheet.getRange("A5:U5").format.rowHeight = 42;
styleBase(tradesSheet, tradesSheet.getUsedRange());

// Trade Legs
const legHeaders = ["Trade no.", "Signal ID", "Entry time (IST)", "Exit time (IST)", "Leg no.", "Symbol", "Side", "Quantity", "Multiplier", "Lots bought", "Total units", "Entry price", "Exit price", "Average exit price", "Price move", "Leg P&L per lot (INR)", "Leg P&L total (INR)", "Target price", "Stop loss", "Expiry (IST)", "Strike price", "Option type", "Instrument type", "Product type", "Underlying", "Exchange", "Exchange token"];
const legRows = [];
trades.forEach((t, idx) => (t.trades || []).forEach((leg, legIdx) => {
  const perLot = round2(legPnlPerLot(leg));
  legRows.push([
    idx + 1, t.signalId, istDate(t.createdAt), istDate(t.closedOn), legIdx + 1, leg.symbol, leg.transactionType,
    Number(leg.quantity || 0), Number(leg.multiplier || 0), Number(t.lotsBought || 0), Number(leg.quantity || 0) * Number(leg.multiplier || 0) * Number(t.lotsBought || 0),
    Number(leg.entryPrice || 0), Number(leg.exitPrice || 0), Number(leg.avgExitPrice || 0), Number(leg.exitPrice || 0) - Number(leg.entryPrice || 0),
    perLot, round2(perLot * Number(t.lotsBought || 0)), Number(leg.targetPrice || 0), Number(leg.stopLoss || 0), istDate(leg.expiry),
    Number(leg.strikePrice || 0), leg.optionType, leg.instrumentType, leg.productType, leg.underlyingSecurity, leg.exchange, leg.exchangeToken,
  ]);
}));
styleTitle(legsSheet, "Option legs", `Each row is one entry/exit leg. Source quantities are preserved; total units include the trade's lots bought.`, legHeaders.length);
legsSheet.getRange(`A5:${excelCol(legHeaders.length)}5`).values = [legHeaders];
legsSheet.getRange(`A6:${excelCol(legHeaders.length)}${5 + legRows.length}`).values = legRows;
styleHeader(legsSheet.getRange(`A5:${excelCol(legHeaders.length)}5`));
legsSheet.tables.add(`A5:${excelCol(legHeaders.length)}${5 + legRows.length}`, true, "TradeLegsTable").style = "TableStyleMedium4";
styleHeader(legsSheet.getRange(`A5:${excelCol(legHeaders.length)}5`));
legsSheet.getRange(`C6:D${5 + legRows.length}`).format.numberFormat = "dd-mmm-yyyy hh:mm";
legsSheet.getRange(`T6:T${5 + legRows.length}`).format.numberFormat = "dd-mmm-yyyy hh:mm";
for (const col of ["L", "M", "N", "O", "P", "Q", "R", "S"]) legsSheet.getRange(`${col}6:${col}${5 + legRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
legsSheet.freezePanes.freezeRows(5);
legsSheet.freezePanes.freezeColumns(2);
legsSheet.getUsedRange().format.autofitColumns();
legsSheet.getRange("B:B").format.columnWidth = 25;
legsSheet.getRange("C:D").format.columnWidth = 21;
legsSheet.getRange("F:F").format.columnWidth = 24;
legsSheet.getRange("A:AA").format.rowHeight = 22;
legsSheet.getRange("A5:AA5").format.rowHeight = 42;
styleBase(legsSheet, legsSheet.getUsedRange());

// Exit details: first table is API exits; second is summary records and their legs.
styleTitle(exitsSheet, "Exit details", "API exit events followed by the strategy's exit-summary leg records.", 12);
const exitHeaders = ["Trade no.", "Signal ID", "Exit event ID", "Closed time (IST)", "Exit value", "P&L per lot (INR)", "P&L %"];
const exitRows = [];
trades.forEach((t, idx) => (t.exits || []).forEach(x => exitRows.push([idx + 1, t.signalId, x._id, istDate(x.closedOn), Number(x.value || 0), Number(x.pnl || 0), Number(x.pnlPercent || 0) / 100])));
exitsSheet.getRange("A5:G5").values = [exitHeaders];
exitsSheet.getRange(`A6:G${5 + exitRows.length}`).values = exitRows;
styleHeader(exitsSheet.getRange("A5:G5"));
exitsSheet.tables.add(`A5:G${5 + exitRows.length}`, true, "ExitEventsTable").style = "TableStyleMedium4";
styleHeader(exitsSheet.getRange("A5:G5"));
exitsSheet.getRange(`D6:D${5 + exitRows.length}`).format.numberFormat = "dd-mmm-yyyy hh:mm";
exitsSheet.getRange(`E6:F${5 + exitRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
exitsSheet.getRange(`G6:G${5 + exitRows.length}`).format.numberFormat = "0.00%";

const summaryStart = 8 + exitRows.length;
exitsSheet.getRange(`A${summaryStart}`).values = [["Exit-summary legs"]];
exitsSheet.getRange(`A${summaryStart}`).format.font = { name: font, size: 12, bold: true, color: colors.title };
const summaryHeaders = ["Trade no.", "Signal ID", "Summary ID", "Title", "Closed time (IST)", "Summary P&L (INR)", "Lots traded", "Summary leg no.", "Summary trade ID", "Symbol", "Quantity (units)", "Exit price"];
const summaryRows = [];
trades.forEach((t, idx) => (t.summary || []).forEach(s => (s.trades || []).forEach((x, legIdx) => summaryRows.push([
  idx + 1, t.signalId, s._id, s.title, istDate(s.closedOn), Number(s.pnl || 0), Number(s.lotsTraded || 0), legIdx + 1, x._id, x.symbol, Number(x.qty || 0), Number(x.price || 0),
]))));
const sh = summaryStart + 2;
exitsSheet.getRange(`A${sh}:L${sh}`).values = [summaryHeaders];
exitsSheet.getRange(`A${sh + 1}:L${sh + summaryRows.length}`).values = summaryRows;
styleHeader(exitsSheet.getRange(`A${sh}:L${sh}`), colors.header2);
exitsSheet.tables.add(`A${sh}:L${sh + summaryRows.length}`, true, "ExitSummaryLegsTable").style = "TableStyleMedium4";
styleHeader(exitsSheet.getRange(`A${sh}:L${sh}`), colors.header2);
exitsSheet.getRange(`E${sh + 1}:E${sh + summaryRows.length}`).format.numberFormat = "dd-mmm-yyyy hh:mm";
exitsSheet.getRange(`F${sh + 1}:F${sh + summaryRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
exitsSheet.getRange(`L${sh + 1}:L${sh + summaryRows.length}`).format.numberFormat = "#,##0.00;[Red](#,##0.00);-";
exitsSheet.freezePanes.freezeRows(5);
exitsSheet.freezePanes.freezeColumns(1);
exitsSheet.getUsedRange().format.autofitColumns();
exitsSheet.getRange("B:C").format.columnWidth = 25;
exitsSheet.getRange("D:D").format.columnWidth = 18;
exitsSheet.getRange("E:E").format.columnWidth = 21;
exitsSheet.getRange("I:I").format.columnWidth = 25;
exitsSheet.getRange("J:J").format.columnWidth = 24;
styleBase(exitsSheet, exitsSheet.getUsedRange());

workbook.recalculate();

const summaryInspect = await workbook.inspect({
  kind: "table",
  range: "Monthly Summary!A2:N20",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 14,
  maxChars: 10000,
});
console.log(summaryInspect.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
  maxChars: 4000,
});
console.log(errors.ndjson);

await fs.mkdir(previewDir, { recursive: true });
for (const sheetName of ["Monthly Summary", "Trades", "Trade Legs", "Exit Details"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  const safe = sheetName.toLowerCase().replace(/\s+/g, "_");
  await fs.writeFile(`${previewDir}/${safe}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(JSON.stringify({ outputPath, tradeCount: trades.length, legCount: legRows.length, exitEventCount: exitRows.length, exitSummaryLegCount: summaryRows.length, totalPnl }, null, 2));



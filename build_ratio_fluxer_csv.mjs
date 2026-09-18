import fs from "node:fs/promises";
import crypto from "node:crypto";

const inputPath = "C:/Users/kavis/OneDrive/Desktop/dhanalgo/ratio_fluxer_trades_raw.json";
const outputPath = "C:/Users/kavis/OneDrive/Desktop/dhanalgo/Ratio-Fluxer_Past_Trades_Detailed_2025-09-18_to_2026-09-18.csv";
const sourceUrl = "https://algos.dhan.co/managers/stratzy/ratio-fluxer-credit-spread-expiry/698075eaf867bf12b20442e5?tab=past-trades";

const payload = JSON.parse(await fs.readFile(inputPath, "utf8"));
if (payload.status !== "success" || !Array.isArray(payload.data)) {
  throw new Error(`Unexpected raw data: ${payload.message || payload.status}`);
}

function istParts(iso) {
  if (!iso) return null;
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
  }).formatToParts(new Date(iso));
  return Object.fromEntries(parts.map((part) => [part.type, part.value]));
}

function formatIst(iso) {
  const p = istParts(iso);
  return p ? `${p.year}-${p.month}-${p.day} ${p.hour}:${p.minute}:${p.second}` : "";
}

function dateIst(iso) {
  const p = istParts(iso);
  return p ? `${p.year}-${p.month}-${p.day}` : "";
}

function monthIst(iso) {
  const p = istParts(iso);
  return p ? `${p.year}-${p.month}` : "";
}

function parseUiEntry(name) {
  const match = String(name || "").match(/:(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})$/);
  if (!match) return { date: "", month: "", datetime: "" };
  const [, day, month, year, hour, minute] = match;
  return {
    date: `${year}-${month}-${day}`,
    month: `${year}-${month}`,
    datetime: `${year}-${month}-${day} ${hour}:${minute}:00`,
  };
}

function round(value, digits = 6) {
  if (value === "" || value === null || value === undefined || Number.isNaN(Number(value))) return "";
  const factor = 10 ** digits;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function oppositeSide(side) {
  return side === "SELL" ? "BUY" : side === "BUY" ? "SELL" : "";
}

function csvValue(value) {
  if (value === null || value === undefined) return "";
  const text = String(value);
  return /[",\r\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

const trades = payload.data.slice().sort((a, b) => new Date(a.createdAt) - new Date(b.createdAt));
const rows = [];

for (let tradeIndex = 0; tradeIndex < trades.length; tradeIndex += 1) {
  const trade = trades[tradeIndex];
  const uiEntry = parseUiEntry(trade.name);
  const legsBySymbol = new Map(trade.trades.map((leg, legIndex) => [leg.symbol, { ...leg, legIndex }]));
  const shortLeg = trade.trades.find((leg) => leg.transactionType === "SELL");
  const longLeg = trade.trades.find((leg) => leg.transactionType === "BUY");
  const strikeWidth = shortLeg && longLeg ? Math.abs(Number(shortLeg.strikePrice) - Number(longLeg.strikePrice)) : "";
  const spreadType = shortLeg?.optionType === "CE" ? "CALL_CREDIT_SPREAD" : shortLeg?.optionType === "PE" ? "PUT_CREDIT_SPREAD" : "";
  const entryNetCredit = shortLeg && longLeg ? Number(shortLeg.entryPrice) - Number(longLeg.entryPrice) : "";
  const createdMs = new Date(trade.createdAt).getTime();
  const closedMs = new Date(trade.closedOn).getTime();

  for (let exitIndex = 0; exitIndex < trade.summary.length; exitIndex += 1) {
    const exit = trade.summary[exitIndex];
    const exitPrices = new Map(exit.trades.map((leg) => [leg.symbol, Number(leg.price)]));
    const exitShortPrice = shortLeg ? exitPrices.get(shortLeg.symbol) : undefined;
    const exitLongPrice = longLeg ? exitPrices.get(longLeg.symbol) : undefined;
    const exitNetDebit = exitShortPrice !== undefined && exitLongPrice !== undefined ? exitShortPrice - exitLongPrice : "";

    for (const exitLeg of exit.trades) {
      const base = legsBySymbol.get(exitLeg.symbol);
      if (!base) throw new Error(`Exit symbol ${exitLeg.symbol} has no matching entry leg for ${trade.signalId}`);
      const direction = base.transactionType === "SELL" ? 1 : -1;
      const eventQty = Number(exitLeg.qty);
      const entryPrice = Number(base.entryPrice);
      const eventExitPrice = Number(exitLeg.price);
      const expiryMs = new Date(base.expiry).getTime();
      const rawEntryIst = formatIst(trade.createdAt);
      const entryMismatchMinutes = uiEntry.datetime
        ? Math.round((new Date(`${rawEntryIst.replace(" ", "T")}+05:30`).getTime() - new Date(`${uiEntry.datetime.replace(" ", "T")}+05:30`).getTime()) / 60000)
        : "";

      rows.push({
        trade_number: tradeIndex + 1,
        entry_month_ui: uiEntry.month,
        close_month_ist: monthIst(trade.closedOn),
        trade_name: trade.name,
        signal_id: trade.signalId,
        entry_time_ui: uiEntry.datetime,
        entry_time_raw_utc: trade.createdAt,
        entry_time_ist_from_raw: rawEntryIst,
        ui_vs_raw_entry_minutes: entryMismatchMinutes,
        trade_closed_time_raw_utc: trade.closedOn,
        trade_closed_time_ist: formatIst(trade.closedOn),
        trade_duration_minutes_raw: Math.round((closedMs - createdMs) / 60000),
        capital_allocated_pct: trade.capitalAllocated,
        lots_bought: trade.lotsBought,
        margin_required: round(trade.marginRequired, 2),
        trade_pnl: round(trade.pnl, 6),
        trade_pnl_pct: round(trade.pnlPercent, 6),
        signal_pnl: round(trade.signalPnl, 6),
        signal_pnl_pct: round(trade.signalPnlPercent, 6),
        exit_event_number: exitIndex + 1,
        exit_event_name: exit.title,
        exit_event_id: exit._id,
        exit_event_time_raw_utc: exit.closedOn,
        exit_event_time_ist: formatIst(exit.closedOn),
        exit_event_lots: exit.lotsTraded,
        exit_event_pnl: round(exit.pnl, 6),
        leg_number: base.legIndex + 1,
        entry_side: base.transactionType,
        exit_side: oppositeSide(base.transactionType),
        symbol: base.symbol,
        underlying: base.underlyingSecurity,
        exchange: base.exchange,
        exchange_token: base.exchangeToken ?? "",
        instrument_type: base.instrumentType,
        product_type: base.productType,
        expiry_raw_utc: base.expiry,
        expiry_date_ist: dateIst(base.expiry),
        expiry_time_ist: formatIst(base.expiry),
        days_to_expiry_from_raw_entry: round((expiryMs - createdMs) / 86400000, 6),
        strike_price: base.strikePrice,
        option_type: base.optionType,
        leg_quantity_lots: base.quantity,
        lot_size_multiplier: base.multiplier,
        trade_leg_total_qty: Number(base.quantity) * Number(base.multiplier) * Number(trade.lotsBought),
        exit_event_qty: eventQty,
        entry_price: round(entryPrice, 6),
        exit_event_price: round(eventExitPrice, 6),
        final_exit_price: round(base.exitPrice, 6),
        avg_exit_price: round(base.avgExitPrice, 6),
        target_price: round(base.targetPrice, 6),
        stop_loss: round(base.stopLoss, 6),
        event_entry_value: round(entryPrice * eventQty, 6),
        event_exit_value: round(eventExitPrice * eventQty, 6),
        event_leg_realized_pnl: round(direction * (entryPrice - eventExitPrice) * eventQty, 6),
        spread_type: spreadType,
        short_strike: shortLeg?.strikePrice ?? "",
        long_strike: longLeg?.strikePrice ?? "",
        strike_width: strikeWidth,
        entry_short_price: shortLeg?.entryPrice ?? "",
        entry_long_price: longLeg?.entryPrice ?? "",
        entry_net_credit: round(entryNetCredit, 6),
        exit_event_short_price: exitShortPrice ?? "",
        exit_event_long_price: exitLongPrice ?? "",
        exit_event_net_debit: round(exitNetDebit, 6),
        source_url: sourceUrl,
      });
    }
  }
}

if (rows.length === 0) throw new Error("No exit-event leg rows were produced");
const headers = Object.keys(rows[0]);
for (const row of rows) {
  if (Object.keys(row).join("|") !== headers.join("|")) throw new Error("Inconsistent CSV columns");
}

const eventGroups = new Map();
for (const row of rows) {
  const key = `${row.signal_id}|${row.exit_event_number}`;
  eventGroups.set(key, (eventGroups.get(key) || 0) + Number(row.event_leg_realized_pnl));
}
for (const trade of trades) {
  trade.summary.forEach((exit, index) => {
    const key = `${trade.signalId}|${index + 1}`;
    const calculated = round(eventGroups.get(key), 2);
    const reported = round(exit.pnl, 2);
    if (Math.abs(calculated - reported) > 1) {
      throw new Error(`Exit P&L mismatch for ${key}: calculated ${calculated}, reported ${reported}`);
    }
  });
}

const csv = `${headers.join(",")}\r\n${rows.map((row) => headers.map((header) => csvValue(row[header])).join(",")).join("\r\n")}\r\n`;
await fs.writeFile(outputPath, `\uFEFF${csv}`, "utf8");
const bytes = await fs.readFile(outputPath);
const digest = crypto.createHash("sha256").update(bytes).digest("hex");

console.log(JSON.stringify({
  outputPath,
  trades: trades.length,
  entryLegs: trades.reduce((sum, trade) => sum + trade.trades.length, 0),
  exitEvents: trades.reduce((sum, trade) => sum + trade.summary.length, 0),
  csvRows: rows.length,
  columns: headers.length,
  firstEntryUi: rows[0].entry_time_ui,
  lastCloseIst: rows.at(-1).trade_closed_time_ist,
  sha256: digest,
}, null, 2));

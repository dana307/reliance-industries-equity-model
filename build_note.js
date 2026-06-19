const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, TabStopType,
  TabStopPosition, LevelFormat
} = require("docx");

// ── Palette ────────────────────────────────────────────────────────────────────
const NAVY   = "1F497D";
const MID_BL = "4472C4";
const LT_BL  = "D9E2F3";
const DARK   = "1A1A2E";
const GRAY   = "595959";
const LT_GRY = "F2F2F2";
const WHITE  = "FFFFFF";
const GREEN  = "375623";
const RED    = "843C0C";

// ── Border helpers ─────────────────────────────────────────────────────────────
const thin  = (c="CCCCCC") => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const thick = (c) => ({ style: BorderStyle.SINGLE, size: 4, color: c });
const noBdr = () => ({ style: BorderStyle.NIL });
const noBdrs = () => ({ top: noBdr(), bottom: noBdr(), left: noBdr(), right: noBdr() });
const thinBdrs = (c="CCCCCC") => ({ top: thin(c), bottom: thin(c), left: thin(c), right: thin(c) });

// ── DXA / margins ─────────────────────────────────────────────────────────────
// 0.75" margins → 1080 DXA each side; content = 12240 - 2160 = 10080 DXA
const CONTENT_W = 10080;
const MARGIN = { top: 1080, right: 1080, bottom: 1080, left: 1080 };

// ── Text helpers ───────────────────────────────────────────────────────────────
const run = (text, opts={}) => new TextRun({ text, font: "Arial", size: opts.size||20,
  bold: opts.bold||false, color: opts.color||"000000", italics: opts.italic||false,
  allCaps: opts.allCaps||false });

const para = (children, opts={}) => new Paragraph({
  children: Array.isArray(children) ? children : [children],
  alignment: opts.align || AlignmentType.LEFT,
  spacing: { before: opts.before||0, after: opts.after||80, line: opts.line||240 },
  indent: opts.indent || undefined,
  border: opts.border || undefined,
  shading: opts.shading || undefined,
  numbering: opts.numbering || undefined,
});

// ── Cell helper ────────────────────────────────────────────────────────────────
const cell = (children, width, opts={}) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  children: Array.isArray(children) ? children : [children],
  shading: opts.bg ? { fill: opts.bg, type: ShadingType.CLEAR } : undefined,
  borders: opts.borders !== undefined ? opts.borders : thinBdrs(),
  verticalAlign: VerticalAlign.CENTER,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  columnSpan: opts.span || 1,
});

const row = (cells) => new TableRow({ children: cells });

// ── Section header bar ─────────────────────────────────────────────────────────
const secHdr = (text) => para(
  [run(text, { bold: true, color: WHITE, size: 19, allCaps: true })],
  { before: 120, after: 40, shading: { fill: NAVY, type: ShadingType.CLEAR },
    indent: { left: 80 } }
);

// ── Small label+value inline ───────────────────────────────────────────────────
const kv = (k, v, vColor="000000") => para([
  run(k + "  ", { color: GRAY, size: 18 }),
  run(v, { bold: true, color: vColor, size: 18 }),
], { after: 30 });

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENT CONTENT
// ══════════════════════════════════════════════════════════════════════════════

// ── Header (every page) ────────────────────────────────────────────────────────
const header = new Header({
  children: [
    para([
      run("RELIANCE INDUSTRIES LTD", { bold: true, color: WHITE, size: 19 }),
      run("\t", {}),
      run("Equity Research  |  India Large-Cap", { color: WHITE, size: 17 }),
    ], {
      align: AlignmentType.LEFT,
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      indent: { left: 60, right: 60 },
      after: 0,
    }),
  ],
});

// ── Footer (every page) ────────────────────────────────────────────────────────
const footer = new Footer({
  children: [
    para([
      run("For academic and portfolio use only. Not investment advice. Data: Screener.in, Company Filings (FY26). ", { color: GRAY, size: 15, italic: true }),
      run("Page ", { color: GRAY, size: 15 }),
      new TextRun({ children: [PageNumber.CURRENT], size: 15, color: GRAY, font: "Arial" }),
    ], {
      before: 60, after: 0,
      border: { top: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
    }),
  ],
});

// ══════════════════════════════════════════════════════════════════════════════
// PAGE 1
// ══════════════════════════════════════════════════════════════════════════════

// ── Title block ────────────────────────────────────────────────────────────────
const titleBlock = [
  // Company / ticker / date row
  para([
    run("Reliance Industries Ltd  ", { bold: true, color: DARK, size: 28 }),
    run("NSE: RELIANCE  |  BSE: 500325", { color: GRAY, size: 19 }),
    run("\t", {}),
    run("June 2026", { color: GRAY, size: 19 }),
  ], { before: 120, after: 60, tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }] }),

  // Metadata bar table: Rating | CMP | DCF Price | Mkt Cap | Sector
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [1600, 1600, 1600, 1780, 3500],
    rows: [
      row([
        cell([para([run("RATING", { size: 16, color: GRAY })], { after: 20 }),
              para([run("NEUTRAL", { bold: true, color: "7F6000", size: 20 })], { after: 0 })],
             1600, { bg: LT_BL }),
        cell([para([run("CMP", { size: 16, color: GRAY })], { after: 20 }),
              para([run("₹1,309.5", { bold: true, size: 20 })], { after: 0 })],
             1600, { bg: LT_BL }),
        cell([para([run("DCF PRICE", { size: 16, color: GRAY })], { after: 20 }),
              para([run("₹381", { bold: true, color: RED, size: 20 })], { after: 0 })],
             1600, { bg: LT_BL }),
        cell([para([run("MKT CAP", { size: 16, color: GRAY })], { after: 20 }),
              para([run("₹17.7 lakh Cr", { bold: true, size: 20 })], { after: 0 })],
             1780, { bg: LT_BL }),
        cell([para([run("Shares: 1,353 Cr  |  FY26 Revenue: ₹10,55,780 Cr  |  FY26 EV: ₹20,30,537 Cr",
                    { size: 17, color: GRAY })], { after: 0 })],
             3500, { bg: LT_GRY }),
      ]),
    ],
  }),
];

// ── Investment Thesis ──────────────────────────────────────────────────────────
const thesis = [
  secHdr("Investment Thesis"),

  para([run(
    "Reliance Industries is India's largest conglomerate by revenue and market cap, operating across " +
    "O2C (Oil-to-Chemicals), Jio Platforms (digital/telecom), Retail (RRVL), and nascent New Energy verticals. " +
    "The stock trades at 21.9x FY26 P/E and 11.3x EV/EBITDA — a discount to telecom peer Bharti Airtel (15.4x) " +
    "and infrastructure peer L&T (17.8x), but a premium to the broader conglomerate basket.",
    { size: 19, color: DARK })
  ], { before: 60, after: 60 }),

  para([run(
    "Our FCFF-based DCF (WACC 11.0%, terminal growth 5%) yields an intrinsic value of ₹381/share — " +
    "a 71% discount to the current market price of ₹1,309. This gap is intentional and expected: " +
    "single-entity DCF cannot capture the embedded optionality in Jio (est. ₹7–9 lakh Cr standalone value), " +
    "RRVL Retail, and the ₹75,000 Cr New Energy platform. The market is effectively paying ₹929/share " +
    "for these high-growth options beyond the hard earnings stream. We rate the stock NEUTRAL pending a " +
    "formal SOTP that closes the valuation gap.",
    { size: 19, color: DARK })
  ], { before: 0, after: 80 }),
];

// ── Financial Summary Table ────────────────────────────────────────────────────
// Cols: Metric | FY24A | FY25A | FY26A | FY25 YoY | FY26 YoY
const FCOLS = [2680, 1480, 1480, 1480, 1480, 1480];  // sum=10080
const hdrRun = (t) => run(t, { bold: true, color: WHITE, size: 18 });
const lblRun = (t) => run(t, { size: 18, color: DARK });
const numRun = (t, bold=false) => run(t, { size: 18, bold, color: DARK });
const chgRun = (t, pos) => run(t, { size: 18, color: pos ? GREEN : (pos===false ? RED : GRAY) });

const finTable = [
  secHdr("Financial Summary  (₹ Crore)"),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: FCOLS,
    rows: [
      // Header
      row([
        cell([para([hdrRun("Metric")],{ after:0 })],           FCOLS[0], { bg: MID_BL }),
        cell([para([hdrRun("FY24A")],{ after:0, align: AlignmentType.RIGHT })], FCOLS[1], { bg: MID_BL }),
        cell([para([hdrRun("FY25A")],{ after:0, align: AlignmentType.RIGHT })], FCOLS[2], { bg: MID_BL }),
        cell([para([hdrRun("FY26A")],{ after:0, align: AlignmentType.RIGHT })], FCOLS[3], { bg: MID_BL }),
        cell([para([hdrRun("YoY FY25")],{ after:0, align: AlignmentType.RIGHT })], FCOLS[4], { bg: MID_BL }),
        cell([para([hdrRun("YoY FY26")],{ after:0, align: AlignmentType.RIGHT })], FCOLS[5], { bg: MID_BL }),
      ]),
      // Revenue
      row([
        cell([para([lblRun("Revenue")])], FCOLS[0]),
        cell([para([numRun("8,99,041")],{align:AlignmentType.RIGHT})],  FCOLS[1]),
        cell([para([numRun("9,62,820")],{align:AlignmentType.RIGHT})],  FCOLS[2]),
        cell([para([numRun("10,55,780",true)],{align:AlignmentType.RIGHT})], FCOLS[3]),
        cell([para([chgRun("+7.1%",true)],{align:AlignmentType.RIGHT})],  FCOLS[4]),
        cell([para([chgRun("+9.7%",true)],{align:AlignmentType.RIGHT})],  FCOLS[5]),
      ]),
      // EBITDA
      row([
        cell([para([lblRun("EBITDA")])], FCOLS[0], { bg: LT_GRY }),
        cell([para([numRun("1,62,498")],{align:AlignmentType.RIGHT})],  FCOLS[1], { bg: LT_GRY }),
        cell([para([numRun("1,65,598")],{align:AlignmentType.RIGHT})],  FCOLS[2], { bg: LT_GRY }),
        cell([para([numRun("1,79,065",true)],{align:AlignmentType.RIGHT})], FCOLS[3], { bg: LT_GRY }),
        cell([para([chgRun("+1.9%",true)],{align:AlignmentType.RIGHT})],  FCOLS[4], { bg: LT_GRY }),
        cell([para([chgRun("+8.1%",true)],{align:AlignmentType.RIGHT})],  FCOLS[5], { bg: LT_GRY }),
      ]),
      // EBITDA Margin
      row([
        cell([para([lblRun("  EBITDA Margin")])], FCOLS[0]),
        cell([para([numRun("18.1%")],{align:AlignmentType.RIGHT})], FCOLS[1]),
        cell([para([numRun("17.2%")],{align:AlignmentType.RIGHT})], FCOLS[2]),
        cell([para([numRun("17.0%",true)],{align:AlignmentType.RIGHT})], FCOLS[3]),
        cell([para([numRun("-90 bps",false)],{align:AlignmentType.RIGHT})], FCOLS[4]),
        cell([para([numRun("-20 bps",false)],{align:AlignmentType.RIGHT})], FCOLS[5]),
      ]),
      // PAT
      row([
        cell([para([lblRun("PAT")])], FCOLS[0], { bg: LT_GRY }),
        cell([para([numRun("69,621")],{align:AlignmentType.RIGHT})],  FCOLS[1], { bg: LT_GRY }),
        cell([para([numRun("69,648")],{align:AlignmentType.RIGHT})],  FCOLS[2], { bg: LT_GRY }),
        cell([para([numRun("80,775",true)],{align:AlignmentType.RIGHT})], FCOLS[3], { bg: LT_GRY }),
        cell([para([chgRun("flat",null)],{align:AlignmentType.RIGHT})],   FCOLS[4], { bg: LT_GRY }),
        cell([para([chgRun("+16.0%",true)],{align:AlignmentType.RIGHT})], FCOLS[5], { bg: LT_GRY }),
      ]),
      // Net Debt
      row([
        cell([para([lblRun("Net Debt")])], FCOLS[0]),
        cell([para([numRun("2,53,494")],{align:AlignmentType.RIGHT})], FCOLS[1]),
        cell([para([numRun("2,67,811")],{align:AlignmentType.RIGHT})], FCOLS[2]),
        cell([para([numRun("2,56,985",true)],{align:AlignmentType.RIGHT})], FCOLS[3]),
        cell([para([numRun("+5.7%")],{align:AlignmentType.RIGHT})], FCOLS[4]),
        cell([para([numRun("-4.0%")],{align:AlignmentType.RIGHT})], FCOLS[5]),
      ]),
    ],
  }),
];

// ── Valuation Summary ──────────────────────────────────────────────────────────
const VCOLS = [2400, 2560, 2560, 2560];  // sum=10080
const valTable = [
  secHdr("Valuation Summary"),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: VCOLS,
    rows: [
      row([
        cell([para([run("Method",{bold:true,color:WHITE,size:18})],{after:0})], VCOLS[0], {bg:MID_BL}),
        cell([para([run("Output",{bold:true,color:WHITE,size:18})],{after:0,align:AlignmentType.RIGHT})], VCOLS[1], {bg:MID_BL}),
        cell([para([run("Assumptions",{bold:true,color:WHITE,size:18})],{after:0})], VCOLS[2], {bg:MID_BL}),
        cell([para([run("vs Market",{bold:true,color:WHITE,size:18})],{after:0,align:AlignmentType.RIGHT})], VCOLS[3], {bg:MID_BL}),
      ]),
      row([
        cell([para([run("DCF (FCFF)",{size:18})])],  VCOLS[0]),
        cell([para([run("₹381 / share",{bold:true,size:18,color:RED})],{align:AlignmentType.RIGHT})], VCOLS[1]),
        cell([para([run("WACC 11.0%, g 5.0%",{size:18,color:GRAY})])], VCOLS[2]),
        cell([para([run("-71%",{bold:true,size:18,color:RED})],{align:AlignmentType.RIGHT})], VCOLS[3]),
      ]),
      row([
        cell([para([run("Market EV/EBITDA",{size:18})])], VCOLS[0], {bg:LT_GRY}),
        cell([para([run("11.3x FY26",{bold:true,size:18})],{align:AlignmentType.RIGHT})], VCOLS[1], {bg:LT_GRY}),
        cell([para([run("Peer median: ~15.4x",{size:18,color:GRAY})])], VCOLS[2], {bg:LT_GRY}),
        cell([para([run("~26% discount",{size:18,color:GREEN})],{align:AlignmentType.RIGHT})], VCOLS[3], {bg:LT_GRY}),
      ]),
      row([
        cell([para([run("Market P/E",{size:18})])], VCOLS[0]),
        cell([para([run("21.9x FY26",{bold:true,size:18})],{align:AlignmentType.RIGHT})], VCOLS[1]),
        cell([para([run("Peer median: ~32.6x",{size:18,color:GRAY})])], VCOLS[2]),
        cell([para([run("~33% discount",{size:18,color:GREEN})],{align:AlignmentType.RIGHT})], VCOLS[3]),
      ]),
    ],
  }),
  para([run(
    "Note: DCF uses single-entity FCFF. Jio Platforms, Retail, and New Energy optionality are not captured " +
    "and account for the ~₹929/share gap between DCF and market price. A full SOTP would be required for a " +
    "formal price target.",
    { size: 16, italic: true, color: GRAY })
  ], { before: 40, after: 60 }),
];

// ══════════════════════════════════════════════════════════════════════════════
// PAGE 2
// ══════════════════════════════════════════════════════════════════════════════
const pageBreak = new Paragraph({ children: [new PageBreak()] });

// ── Comps table ────────────────────────────────────────────────────────────────
// 8 columns: Company | Mkt Cap | EV | Revenue | EBITDA | EBITDA % | EV/EBITDA | P/E
const CCOLS = [2100, 1200, 1200, 1280, 1280, 1120, 1100, 800]; // sum=10080
const chdr = (t) => run(t,{bold:true,color:WHITE,size:17});
const cdat = (t,b=false,c="000000") => run(t,{size:17,bold:b,color:c});

const compsTable = [
  secHdr("Comparable Company Analysis  (FY25, ₹ Crore)"),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: CCOLS,
    rows: [
      row([
        cell([para([chdr("Company")],{after:0})],           CCOLS[0],{bg:MID_BL}),
        cell([para([chdr("Mkt Cap")],{after:0,align:AlignmentType.RIGHT})],  CCOLS[1],{bg:MID_BL}),
        cell([para([chdr("EV")],{after:0,align:AlignmentType.RIGHT})],       CCOLS[2],{bg:MID_BL}),
        cell([para([chdr("Revenue")],{after:0,align:AlignmentType.RIGHT})],  CCOLS[3],{bg:MID_BL}),
        cell([para([chdr("EBITDA")],{after:0,align:AlignmentType.RIGHT})],   CCOLS[4],{bg:MID_BL}),
        cell([para([chdr("EBT Mg%")],{after:0,align:AlignmentType.RIGHT})],  CCOLS[5],{bg:MID_BL}),
        cell([para([chdr("EV/EBTDA")],{after:0,align:AlignmentType.RIGHT})], CCOLS[6],{bg:MID_BL}),
        cell([para([chdr("P/E")],{after:0,align:AlignmentType.RIGHT})],      CCOLS[7],{bg:MID_BL}),
      ]),
      // RIL FY26 (bold, highlighted)
      row([
        cell([para([cdat("Reliance (FY26)",true)])],                           CCOLS[0],{bg:LT_BL}),
        cell([para([cdat("17,72,553",true)],{align:AlignmentType.RIGHT})],    CCOLS[1],{bg:LT_BL}),
        cell([para([cdat("20,29,537",true)],{align:AlignmentType.RIGHT})],    CCOLS[2],{bg:LT_BL}),
        cell([para([cdat("10,55,780",true)],{align:AlignmentType.RIGHT})],    CCOLS[3],{bg:LT_BL}),
        cell([para([cdat("1,79,065",true)],{align:AlignmentType.RIGHT})],     CCOLS[4],{bg:LT_BL}),
        cell([para([cdat("17.0%",true)],{align:AlignmentType.RIGHT})],        CCOLS[5],{bg:LT_BL}),
        cell([para([cdat("11.3x",true)],{align:AlignmentType.RIGHT})],        CCOLS[6],{bg:LT_BL}),
        cell([para([cdat("21.9x",true)],{align:AlignmentType.RIGHT})],        CCOLS[7],{bg:LT_BL}),
      ]),
      // Peers
      ...[
        ["Adani Enterprises",  "2,63,909",  "3,47,728",  "97,895",   "14,252", "14.6%", "24.4x", "71.6x"],
        ["Tata Motors",        "2,66,787",  "3,03,327",  "4,39,695", "56,138", "12.8%",  "5.4x", "11.2x"],
        ["Larsen & Toubro",    "5,35,675",  "6,13,279",  "2,55,734", "34,427", "13.5%", "17.8x", "32.6x"],
        ["Bharti Airtel",     "11,13,302", "13,08,944",  "1,72,985", "85,060", "49.2%", "15.4x", "36.3x"],
        ["ITC",                "3,55,462",  "3,21,027",   "75,323",  "25,839", "34.3%", "12.4x", "17.0x"],
      ].map(([co,mc,ev,rev,ebt,mg,eveb,pe],i) =>
        row([
          cell([para([cdat(co)])],               CCOLS[0], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(mc)],{align:AlignmentType.RIGHT})],  CCOLS[1], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(ev)],{align:AlignmentType.RIGHT})],  CCOLS[2], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(rev)],{align:AlignmentType.RIGHT})], CCOLS[3], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(ebt)],{align:AlignmentType.RIGHT})], CCOLS[4], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(mg)],{align:AlignmentType.RIGHT})],  CCOLS[5], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(eveb)],{align:AlignmentType.RIGHT})],CCOLS[6], i%2===1?{bg:LT_GRY}:{}),
          cell([para([cdat(pe)],{align:AlignmentType.RIGHT})],  CCOLS[7], i%2===1?{bg:LT_GRY}:{}),
        ])
      ),
      // Peer median
      row([
        cell([para([cdat("Peer Median",true,"4472C4")])], CCOLS[0]),
        cell([para([cdat("—",true)],{align:AlignmentType.RIGHT})], CCOLS[1]),
        cell([para([cdat("—",true)],{align:AlignmentType.RIGHT})], CCOLS[2]),
        cell([para([cdat("—",true)],{align:AlignmentType.RIGHT})], CCOLS[3]),
        cell([para([cdat("—",true)],{align:AlignmentType.RIGHT})], CCOLS[4]),
        cell([para([cdat("24.9%",true)],{align:AlignmentType.RIGHT})], CCOLS[5]),
        cell([para([cdat("15.4x",true,"4472C4")],{align:AlignmentType.RIGHT})], CCOLS[6]),
        cell([para([cdat("32.6x",true,"4472C4")],{align:AlignmentType.RIGHT})], CCOLS[7]),
      ]),
    ],
  }),
  para([run(
    "RIL trades at a 26% discount on EV/EBITDA and 33% on P/E vs peer medians, reflecting heavier capital intensity " +
    "and near-term FCFF compression from ongoing capex. Bharti Airtel (15.4x EV/EBITDA) is the closest telecom/infra comparable for Jio.",
    { size: 16, italic: true, color: GRAY })
  ], { before: 40, after: 60 }),
];

// ── Scenario Analysis ──────────────────────────────────────────────────────────
// 3-col table: Bull / Base / Bear (side by side)
const SCOLS = [1500, 2860, 2860, 2860]; // sum=10080
const scenarios = [
  secHdr("Scenario Analysis"),
  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: SCOLS,
    rows: [
      row([
        cell([para([run("",{size:18})])], SCOLS[0]),
        cell([para([run("BULL CASE",{bold:true,color:GREEN,size:19})],{after:0,align:AlignmentType.CENTER})], SCOLS[1]),
        cell([para([run("BASE CASE",{bold:true,color:"7F6000",size:19})],{after:0,align:AlignmentType.CENTER})], SCOLS[2]),
        cell([para([run("BEAR CASE",{bold:true,color:RED,size:19})],{after:0,align:AlignmentType.CENTER})], SCOLS[3]),
      ]),
      row([
        cell([para([run("Probability",{size:17,bold:true})])], SCOLS[0],{bg:LT_GRY}),
        cell([para([run("30%",{size:17,bold:true,color:GREEN})],{align:AlignmentType.CENTER})], SCOLS[1],{bg:LT_GRY}),
        cell([para([run("50%",{size:17,bold:true,color:"7F6000"})],{align:AlignmentType.CENTER})], SCOLS[2],{bg:LT_GRY}),
        cell([para([run("20%",{size:17,bold:true,color:RED})],{align:AlignmentType.CENTER})], SCOLS[3],{bg:LT_GRY}),
      ]),
      row([
        cell([para([run("Jio Value",{size:17,bold:true})])], SCOLS[0]),
        cell([para([run("Successful monetisation of 5G; Jio platform IPO at ₹10L Cr valuation drives re-rating",{size:17})],{align:AlignmentType.CENTER})], SCOLS[1]),
        cell([para([run("Steady ARPU uplift; Jio contributes 40%+ of consol. EBITDA by FY28",{size:17})],{align:AlignmentType.CENTER})], SCOLS[2]),
        cell([para([run("Intense competition from Airtel/Vi delays ARPU recovery; Jio IPO deferred",{size:17})],{align:AlignmentType.CENTER})], SCOLS[3]),
      ]),
      row([
        cell([para([run("New Energy",{size:17,bold:true})])], SCOLS[0],{bg:LT_GRY}),
        cell([para([run("₹75K Cr capex delivers giga-scale by FY28; green H₂ export contracts materialise",{size:17})],{align:AlignmentType.CENTER})], SCOLS[1],{bg:LT_GRY}),
        cell([para([run("Projects on track; contribution to EBITDA begins FY27 at modest scale",{size:17})],{align:AlignmentType.CENTER})], SCOLS[2],{bg:LT_GRY}),
        cell([para([run("Policy delays; solar manufacturing margin disappoints; capex overruns widen net debt",{size:17})],{align:AlignmentType.CENTER})], SCOLS[3],{bg:LT_GRY}),
      ]),
      row([
        cell([para([run("O2C Cycle",{size:17,bold:true})])], SCOLS[0]),
        cell([para([run("Crude normalises; refining spreads recover to $9–10/bbl; petchem demand returns",{size:17})],{align:AlignmentType.CENTER})], SCOLS[1]),
        cell([para([run("O2C EBITDA holds at ₹60–65K Cr; margins stable at ~10%",{size:17})],{align:AlignmentType.CENTER})], SCOLS[2]),
        cell([para([run("Refining margin compression; China petchem oversupply persists into FY27",{size:17})],{align:AlignmentType.CENTER})], SCOLS[3]),
      ]),
      row([
        cell([para([run("Implied Price",{size:18,bold:true})])], SCOLS[0],{bg:LT_GRY}),
        cell([para([run("₹1,700–2,000+",{size:19,bold:true,color:GREEN})],{align:AlignmentType.CENTER})], SCOLS[1],{bg:LT_GRY}),
        cell([para([run("₹1,100–1,400",{size:19,bold:true,color:"7F6000"})],{align:AlignmentType.CENTER})], SCOLS[2],{bg:LT_GRY}),
        cell([para([run("₹800–950",{size:19,bold:true,color:RED})],{align:AlignmentType.CENTER})], SCOLS[3],{bg:LT_GRY}),
      ]),
    ],
  }),
];

// ── Key Risks ──────────────────────────────────────────────────────────────────
const risks = [
  secHdr("Key Risks"),

  new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [200, 4940, 200, 4740],
    rows: [
      row([
        cell([para([run("▪",{color:RED,bold:true,size:18})])], 200, {borders:noBdrs()}),
        cell([para([
          run("Refining margin risk: ",{bold:true,size:18}),
          run("GRM compression from Chinese overcapacity and weak global petchem demand could weigh on O2C segment (est. 50% of EBITDA).",{size:18})
        ])], 4940, {borders:noBdrs()}),
        cell([para([run("▪",{color:RED,bold:true,size:18})])], 200, {borders:noBdrs()}),
        cell([para([
          run("Leverage risk: ",{bold:true,size:18}),
          run("Net debt of ₹2.57 lakh Cr remains elevated; continued capex for New Energy could delay FCF inflection.",{size:18})
        ])], 4740, {borders:noBdrs()}),
      ]),
      row([
        cell([para([run("▪",{color:RED,bold:true,size:18})])], 200, {borders:noBdrs()}),
        cell([para([
          run("Regulatory risk: ",{bold:true,size:18}),
          run("Telecom tariff regulation, AGR-style levies, or data localisation mandates could impair Jio's economics.",{size:18})
        ])], 4940, {borders:noBdrs()}),
        cell([para([run("▪",{color:RED,bold:true,size:18})])], 200, {borders:noBdrs()}),
        cell([para([
          run("Conglomerate discount: ",{bold:true,size:18}),
          run("Market may continue to apply a discount for complexity, related-party transactions, and governance perceptions.",{size:18})
        ])], 4740, {borders:noBdrs()}),
      ]),
    ],
  }),
];

// ══════════════════════════════════════════════════════════════════════════════
// ASSEMBLE DOCUMENT
// ══════════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: { config: [] },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: MARGIN,
      },
    },
    headers: { default: header },
    footers: { default: footer },
    children: [
      ...titleBlock,
      ...thesis,
      ...finTable,
      ...valTable,
      pageBreak,
      ...compsTable,
      ...scenarios,
      ...risks,
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("RIL_Investment_Note.docx", buf);
  console.log("Written: RIL_Investment_Note.docx");
});

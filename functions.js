/*
 * PUMPCALC Custom Functions for Excel (Office.js)
 * ---------------------------------------------------------------
 * Auto-assembled from src/core.js (verified against VBA ground truth,
 * 496/496 test cases passing - see verify.js / verify2.js).
 * DO NOT hand-edit the calculation logic here; edit src/core.js and
 * re-run the assembly step (see README.md) instead.
 * ---------------------------------------------------------------
 */

const pipeIdTable = {"schCols": ["STD", "10", "20", "30", "40", "60", "80", "100", "120", "140", "160"], "rows": [{"normID": 0.25, "vals": [0.364, null, null, 0.41, null, null, null, null, null, null, null]}, {"normID": 0.5, "vals": [0.622, null, null, 0.674, 0.622, null, 0.546, null, 0.466, null, 0.446]}, {"normID": 0.75, "vals": [0.824, null, null, 0.884, 0.824, null, 0.742, null, 0.612, null, 0.612]}, {"normID": 1, "vals": [1.049, null, 1.097, 1.097, 1.049, null, 0.957, null, 0.815, null, 0.815]}, {"normID": 1.5, "vals": [1.61, null, 1.682, 1.682, 1.61, null, 1.5, null, 1.388, null, 1.388]}, {"normID": 2, "vals": [2.067, null, 2.157, 2.157, 2.067, null, 1.939, null, 1.687, null, 1.687]}, {"normID": 2.5, "vals": [2.469, null, null, null, 2.469, null, 2.323, null, null, null, 2.125]}, {"normID": 3, "vals": [3.068, null, 3.26, 3.26, 3.068, null, 2.9, null, 2.624, null, 2.624]}, {"normID": 4, "vals": [4.026, null, 4.26, 4.26, 4.026, null, 3.826, null, 3.624, null, 3.438]}, {"normID": 6, "vals": [6.065, null, 6.357, 6.357, 6.065, null, 5.761, null, 5.501, null, 5.187]}, {"normID": 8, "vals": [7.981, null, 8.125, 8.071, 7.981, 7.813, 7.625, 7.437, 7.187, 7.001, 6.813]}, {"normID": 10, "vals": [10.02, null, 10.25, 10.136, 10.02, 9.75, 9.562, 9.312, 9.062, 8.75, 8.5]}, {"normID": 12, "vals": [12, null, 12.25, 12.09, 11.938, 11.626, 11.374, 11.062, 10.75, 10.5, 10.126]}, {"normID": 14, "vals": [13.25, 13.5, 13.376, 13.25, 13.124, 12.812, 12.5, 12.124, 11.812, 11.5, 11.188]}, {"normID": 16, "vals": [15.25, 15.5, 15.376, 15.25, 15, 14.688, 14.312, 13.938, 13.562, 13.124, 12.812]}, {"normID": 18, "vals": [17.25, 17.5, 17.376, 17.124, 16.876, 16.5, 16.124, 15.688, 15.25, 14.876, 14.438]}, {"normID": 20, "vals": [19.25, 19.5, 19.25, 19, 18.812, 18.376, 17.938, 17.438, 17, 16.5, 16.062]}, {"normID": 22, "vals": [21.25, 21.5, 21.25, null, 21, 20.25, 19.75, 19.25, 18.75, 18.25, 17.75]}, {"normID": 24, "vals": [23.25, 23.5, 23.25, 22.876, 22.624, 22.062, 21.562, 20.938, 20.376, 19.876, 19.312]}, {"normID": 26, "vals": [25.25, 25.376, 25, null, null, null, null, null, null, null, null]}, {"normID": 28, "vals": [27.25, 27.376, 27, 26.75, null, null, null, null, null, null, null]}, {"normID": 30, "vals": [29.25, 29.376, 29, 28.75, null, null, null, null, null, null, null]}, {"normID": 32, "vals": [31.25, 31.376, 31, 30.75, 30.624, null, null, null, null, null, null]}, {"normID": 34, "vals": [33.25, 33.312, 33, 32.75, 32.624, null, null, null, null, null, null]}, {"normID": 36, "vals": [35.25, 35.376, 35, 34.75, 34.5, null, null, null, null, null, null]}, {"normID": 42, "vals": [null, null, 41, 40.75, 40.5, null, null, null, null, null, null]}, {"normID": 48, "vals": [null, 48, 48, 48, 48, null, null, null, null, null, null]}, {"normID": 54, "vals": [null, 54, 54, 54, 54, 54, 54, 54, 54, 54, 54]}, {"normID": 60, "vals": [null, 60, 60, 60, 60, 60, 60, 60, 60, 60, 60]}, {"normID": 66, "vals": [null, 66, 66, 66, 66, 66, 66, 66, 66, 66, 66]}, {"normID": 72, "vals": [null, 72, 72, 72, 72, 72, 72, 72, 72, 72, 72]}, {"normID": 84, "vals": [null, 84, 84, 84, 84, 84, 84, 84, 84, 84, 84]}, {"normID": 96, "vals": [null, 96, 96, 96, 96, 96, 96, 96, 96, 96, 96]}, {"normID": 108, "vals": [null, 108, 108, 108, 108, 108, 108, 108, 108, 108, 108]}, {"normID": 120, "vals": [null, 120, 120, 120, 120, 120, 120, 120, 120, 120, 120]}]};

/*
 * PUMPCALC core calculation library
 * -----------------------------------------------------------------------
 * Ported 1:1 from Module_For_Equation.bas (original VBA) / lambda_defs.py
 * (verified Excel LAMBDA rewrite used in the desktop WebFix builds).
 * This file has NO dependency on Office.js - it is plain JS so it can be
 * unit tested with plain `node` before being wired into Custom Functions.
 * -----------------------------------------------------------------------
 */

/* pipeIdTable injected below */

// ---- small helpers -------------------------------------------------------
function ln(x) { return Math.log(x); }
function log10(x) { return Math.log(x) / Math.LN10; }

// ---- Calc_J / Calc_JA / Calc_RA ------------------------------------------
// VBA original used a Do-While Newton-Raphson loop; the verified Excel port
// (lambda_defs.py) replaced that with a fixed 1000-iteration fold (REDUCE)
// after confirming <1e-6 relative error across 20,000 random trials. We
// reuse the same fixed-iteration approach here for parity with the xlsm.
function Calc_J(ValU, RAAss, ValAS, ValAP) {
  return ValU === 1 ? RAAss * ValAS * ValAP : ValAS * ValAP + 2 * ValAS * ln(RAAss) + 1 - RAAss ** 2;
}

function Calc_JA(ValU, RAAss, ValAS) {
  return ValU === 2 ? (2 * ValAS) / RAAss - 2 * RAAss : 0;
}

function Calc_RA(ValU, ValAS, ValAP) {
  let oldRA = 1;
  for (let i = 0; i < 1000; i++) {
    const ValJ = Calc_J(ValU, oldRA, ValAS, ValAP);
    const ValJA = Calc_JA(ValU, oldRA, ValAS);
    oldRA = ValU === 1 ? ValJ + (ValJ * ln(oldRA)) / ValAP + 1 : (-1 * ValJ) / ValJA + oldRA;
  }
  return oldRA;
}

// ---- Pipe_ID ---------------------------------------------------------------
// Exact-match lookup against a baked-in snapshot of the Pipe_ID sheet
// (schCols x normID rows). Schedule ("Sch") is compared as text so both
// numeric (40, 80, ...) and text ("STD") schedule values from the sheet
// match correctly - this was the source of an earlier #VALUE! bug where
// functions.json declared "sch" as strict type "string", which made Office
// reject numeric schedule cells before this function ever ran; "sch" is now
// declared "any" in functions.json/JSDoc and String(Sch) here normalizes it.
// If either match fails, falls back to the first row/col, mirroring the
// original VBA Do-Until search degrading past the end of its data.
//
// NOTE: this table is a static snapshot, not a live read of the workbook's
// Pipe_ID sheet. Adding a new pipe size or changing a schedule value on the
// Pipe_ID sheet has NO effect on this function - the pipeIdTable object
// below must be edited (and the add-in redeployed to GitHub Pages) for
// that change to take effect. This trade-off was chosen deliberately over
// a live-read implementation (which would require a shared runtime and a
// larger manifest/project restructure) to keep the add-in simple and
// reliable.
function Pipe_ID(NormID, Sch) {
  const schKey = String(Sch).trim();
  let rowIdx = pipeIdTable.rows.findIndex((r) => r.normID === NormID);
  if (rowIdx === -1) rowIdx = 0;
  let colIdx = pipeIdTable.schCols.findIndex((c) => c === schKey);
  if (colIdx === -1) colIdx = 0;
  const val = pipeIdTable.rows[rowIdx].vals[colIdx];
  return val === null || val === undefined ? 0 : val;
}

// ---- two-phase / fluid property helpers -----------------------------------
function VaporVolFrac(WtLiqM, WtVapM, DenLiqM, DenVapM) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const QL = Wliq / DenLiq / 3600;
  const QV = Wvap / DenVap / 3600;
  return QV / (QV + QL);
}

function AvgDensity(WtLiqM, WtVapM, DenLiqM, DenVapM) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const Wtot = Wliq + Wvap;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const DenHomo = Wtot / (Wliq / DenLiq + Wvap / DenVap);
  return DenHomo / 0.0624;
}

function DPHomogeneous(WtLiqM, WtVapM, DenLiqM, DenVapM, e, PipeIDVal) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const Wtot = Wliq + Wvap;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const DenHomo = Wtot / (Wliq / DenLiq + Wvap / DenVap);
  const ID = PipeIDVal / 12;
  const Re = (6.316 * Wtot) / (ID * 12) / 0.1;
  const a = -2 * log10(e / ID / 3.7 + 12 / Re);
  const b = -2 * log10(e / ID / 3.7 + (2.51 * a) / Re);
  const c = -2 * log10(e / ID / 3.7 + (2.51 * b) / Re);
  const f = (a - (b - a) ** 2 / (c - 2 * b + a)) ** -2;
  return (f * 0.000336 * Wtot ** 2) / DenHomo / (ID * 12) ** 5 / 4.42075824;
}

// shared Dukler-correlation RV solver (fixed 150-iteration fold, matches the
// verified Excel REDUCE() port - originally a self-recursive VBA loop)
function dukler_RV(Wliq, Wvap, DenLiq, DenVap, VisLiq, VisVap, Diam, AREA) {
  const QL = Wliq / DenLiq / 3600;
  const QV = Wvap / DenVap / 3600;
  const F_v = QV / (QV + QL);
  const F_l = 1 - F_v;
  const Vm = (QV + QL) / AREA;
  const NFR = Vm ** 2 / 32.174 / Diam;
  const x = Wvap / (Wliq + Wvap);
  const GT = (Wliq + Wvap) / AREA;

  let Rv = 0.5;
  for (let i = 0; i < 150; i++) {
    const NRe1 = (Diam * GT) / (2.42 * (Rv * VisVap + (1 - Rv) * VisLiq));
    const Z = (NRe1 ** (1 / 6) * NFR ** (1 / 8)) / F_l ** 0.25;
    const Charlie =
      Z < 10
        ? -0.16367 + 0.31037 * Z - 0.03525 * Z ** 2 + 0.001366 * Z ** 3
        : 0.75545 + 0.003585 * Z - 0.1436e-4 * Z ** 2;
    Rv = Charlie / ((1 / x - 1) * (DenVap / DenLiq) + 1);
  }
  return { RV: Rv, F_l, GT, NFR, x };
}

function DPDukler(WtLiqM, WtVapM, DenLiqM, DenVapM, VisLiq, VisVap, PipeIDVal, e) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const Wtot = Wliq + Wvap;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;

  const { RV, F_l, GT } = dukler_RV(Wliq, Wvap, DenLiq, DenVap, VisLiq, VisVap, Diam, AREA);

  const LF_l = ln(F_l);
  const H = 1 - LF_l / (1.281 + 0.478 * LF_l + 0.444 * LF_l ** 2 + 0.094 * LF_l ** 3 + 0.00843 * LF_l ** 4);
  const DenNS = DenLiq * F_l + DenVap * (1 - F_l);
  const VisNS = (VisLiq * F_l + VisVap * (1 - F_l)) * 2.42;
  const Bravo1 = (DenLiq / DenNS) * (F_l ** 2 / (1 - RV)) + (DenVap / DenNS) * ((1 - F_l) ** 2 / RV);
  const ReTP1 = (4 * Wtot * Bravo1) / (3.14 * Diam * VisNS);

  let fcal = 0.001;
  for (let i = 0; i < 150; i++) {
    fcal = 1 / (-2 * log10(e / Diam / 3.7 + 2.51 / ReTP1 / fcal)) ** 2;
  }
  const f = fcal / 4;
  // NOTE: original VBA/LAMBDA uses GT (= Wtot/AREA), not Wtot, in this final term.
  return (2 * ((GT / 3600) ** 2) * f * 100 * Bravo1 * H) / (32.2 * Diam * DenNS * 144) / 4.42075824;
}

function InplaceDensity(WtLiqM, WtVapM, DenLiqM, DenVapM, VisLiq, VisVap, PipeIDVal) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;

  const { RV } = dukler_RV(Wliq, Wvap, DenLiq, DenVap, VisLiq, VisVap, Diam, AREA);
  return (DenVap * RV + DenLiq * (1 - RV)) / 0.0624;
}

function TwoPhaseVelocity(WtLiqM, WtVapM, DenLiqM, DenVapM, PipeIDVal) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;
  const QL = Wliq / DenLiq / 3600;
  const QV = Wvap / DenVap / 3600;
  return ((QV + QL) / AREA) * 0.3048;
}

// ---- Baker flow-pattern map -------------------------------------------------
function BakerXval(WtLiqM, WtVapM, DenLiqM, DenVapM, MWl, SurTen, VisLiq, PipeIDVal) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;
  const ST = SurTen === 0 ? (((57 + 2.3 * MWl) / MWl) * ((DenLiq - DenVap) / 62.3)) ** 4 : SurTen;
  const La = ((DenVap / 0.075) * (DenLiq / 62.3)) ** 0.5;
  const P = (73 / ST) * (VisLiq * (62.3 / DenLiq) ** 2) ** (1 / 3);
  const G = Wvap / AREA;
  const L = Wliq / AREA;
  return (L * P * La) / G;
}

function BakerYval(WtVapM, DenLiqM, DenVapM, PipeIDVal) {
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;
  const La = ((DenVap / 0.075) * (DenLiq / 62.3)) ** 0.5;
  const G = Wvap / AREA;
  return G / La;
}

function Baker(X1, Y1) {
  if (X1 < 0.1 || X1 > 10000 || Y1 < 100 || Y1 > 100000) return 'Out of Range';
  const x = (log10(X1) - log10(0.1)) / (log10(10000) - log10(0.1));
  const Y = (log10(Y1) - log10(100)) / (log10(100000) - log10(100));
  const Line1 = 0.19015 * x ** 3 - 0.38395 * x ** 2 - 0.91636 * x + 0.94269;
  const Line2 = 0.41754 * x ** 3 - 0.99032 * x ** 2 + 0.04973 * x + 0.61096;
  const line3 = -3.59472 * x ** 3 + 5.82745 * x ** 2 - 3.13262 * x + 1.37817;
  const line4 = 7.9972 * x ** 3 - 7.0625 * x ** 2 + 1.35609 * x + 0.65341;
  const line5 = -0.36 * x + 0.535;
  const line6 = -133.333 * x ** 2 + 173 * x - 55.10667;
  const line7 = -9.18737 * x ** 3 + 23.70631 * x ** 2 - 21.7587 * x + 7.33735;

  if (x <= 0.294142932210028) {
    if (Y <= Line2) return 'Stratified';
    if (Y <= Line1) return 'Wave';
    if (Y <= line3) return 'Annular';
    return 'Dispersed';
  }
  if (x <= 0.448315796039707) {
    if (Y <= Line2) return 'Stratified';
    if (Y <= Line1) return 'Wave';
    if (Y <= line4) return 'Slug';
    if (Y <= line3) return 'Annular';
    return 'Dispersed';
  }
  if (x <= 0.571278202693418) {
    if (Y <= Line1) return 'Stratified';
    if (Y <= line4) return 'Slug';
    if (Y <= line3) return 'Annular';
    return 'Dispersed';
  }
  if (x <= 0.609274283833812) {
    if (Y <= Line1) return 'Stratified';
    if (Y <= line5) return 'Plug';
    if (Y <= line4) return 'Slug';
    if (Y <= line3) return 'Annular';
    return 'Dispersed';
  }
  if (x <= 0.637634634557965) {
    if (Y <= Line1) return 'Stratified';
    if (Y <= line5) return 'Plug';
    if (Y <= line4) return 'Slug';
    if (Y <= line7) return 'Annular';
    if (Y <= line6) return 'Bubble or Froth';
    return 'Dispersed';
  }
  if (x <= 0.943008707285949) {
    if (Y <= Line1) return 'Stratified';
    if (Y <= line5) return 'Plug';
    if (Y <= line7) return 'Slug';
    return 'Bubble or Froth';
  }
  return Y <= line7 ? 'Plug' : 'Bubble or Froth';
}

// ---- Griffith-Wallis flow-pattern map --------------------------------------
function GriffithWallisXval(WtLiqM, WtVapM, DenLiqM, DenVapM, PipeIDVal) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const Diam = PipeIDVal / 12;
  const AREA = (Diam ** 2 * 3.14) / 4;
  const QL = Wliq / DenLiq / 3600;
  const QV = Wvap / DenVap / 3600;
  const Vm = (QV + QL) / AREA;
  return Vm ** 2 / 32.174 / Diam;
}

function GriffithWallisYval(WtLiqM, WtVapM, DenLiqM, DenVapM) {
  const Wliq = WtLiqM * 2.2046;
  const Wvap = WtVapM * 2.2046;
  const DenLiq = DenLiqM * 0.0624;
  const DenVap = DenVapM * 0.0624;
  const QL = Wliq / DenLiq / 3600;
  const QV = Wvap / DenVap / 3600;
  return QV / (QV + QL);
}

function Griffith(x, Y) {
  if (x < 0 || x > 400 || Y < 0 || Y > 1) return 'Out of Range';
  const Line1 = 0.06903 * x ** -0.650874;
  const Line2 = -0.00037612839 * x + 0.20009027081;
  const line3 = -0.0000034677 * x ** 3 + 0.0003366929 * x ** 2 - 0.0173144767 * x + 1.1758057232;
  if (x <= 0.24) return Y <= Line1 ? 'Bubble' : 'Slug';
  if (x <= 80) {
    if (Y <= Line2) return 'Bubble';
    if (Y <= line3) return 'Slug';
    return 'Annular Mist';
  }
  return 'Annular Mist';
}

// ---- head/pressure conversion & pump sizing --------------------------------
function CalcPres(Head, SG) {
  return (Head * SG) / 10;
}
function CalcHead(Press, SG) {
  return (Press * 10) / SG;
}

function CalcPumpEff(FlowRate) {
  const GPM = 4.4029 * FlowRate;
  if (GPM < 10) return 6.9;
  return (-1.1237557 * ln(GPM) + 20.3592672) * ln(GPM) - 0.79087 - 76.536328 / ln(GPM);
}

function CalcPumpEffAnother(DesRate, SpGr, DiffPress) {
  let Q = 4.4029 * DesRate;
  const H = ((DiffPress * 10) / SpGr) * 3.28084;
  if (Q > 1000) return null; // VBA: Exit Function (no return value)
  if (Q < 25) return null;
  let RealQ;
  if (Q < 100) {
    RealQ = Q;
    Q = 100;
  } else {
    RealQ = 200;
  }
  const a = 80;
  const b = -0.2855 * H;
  const c = 0.000378 * Q * H;
  const D = -0.000000238 * Q ** 2 * H;
  const e = 0.000539 * H ** 2;
  const f = -0.000000639 * Q * H ** 2;
  const G = 0.0000000004 * Q ** 2 * H ** 2;
  let result = a + b + c + D + e + f + G;
  if (RealQ < 100) result -= 0.35 * (100 - RealQ);
  return result;
}

function SelMotorPower(DesRate, SpGr, DiffPress) {
  const GPM = 4.4029 * DesRate;
  const PumpEff = GPM < 10 ? 6.9 : (-1.1237557 * ln(GPM) + 20.3592672) * ln(GPM) - 0.79087 - 76.536328 / ln(GPM);
  const HydPower = (DesRate * DiffPress) / 36.73;
  const BrkPower = HydPower / (PumpEff / 100);
  const Y = BrkPower / 0.746;
  const MotEff = (-0.30546097 * ln(Y) + 5.323494) * ln(Y) + 72;
  const MotPower = BrkPower / (MotEff / 100);
  const Z = MotPower;

  const steps = [0.4, 0.75, 1.5, 2.2, 3.7, 5.5, 7.5, 11, 15, 19, 22, 30, 37, 45, 55, 75, 110, 150];
  for (const s of steps) {
    if (Z <= s) return s;
  }
  return Z;
}

function EstNPSHR(MetricQ) {
  const Q = MetricQ * 4.403;
  const val = (Q ** 0.67 * 10 ** -0.612) / 3.281 + 0.6;
  return val < 1.83 ? 1.83 : val;
}

// ---- EqLenFactor ------------------------------------------------------------
// a = pipe ID (inch), b = straight length (m). Piecewise-linear interpolation
// against 17 breakpoints, ported verbatim from Module_For_Equation.bas.
function EqLenFactor(a, b) {
  const x = b;
  const Temp = a;
  if (!(Temp > 0)) return null;

  const f = new Array(18); // 1-indexed, f[0] unused
  f[1] = 0.00001 * x ** 2 + 0.0275 * x - 0.066;
  f[2] = 0.00002 * x ** 2 + 0.0448 * x + 0.04;
  f[3] = 0.00005 * x ** 2 + 0.0602 * x + 0.2518;
  f[4] = 0.0001 * x ** 2 + 0.0895 * x + 0.5985;
  f[5] = 0.0002 * x ** 2 + 0.144 * x + 0.6265;
  f[6] = 0.0003 * x ** 2 + 0.1901 * x + 1.0593;
  f[7] = 0.0005 * x ** 2 + 0.2092 * x + 1.8022;
  f[8] = 0.0007 * x ** 2 + 0.2659 * x + 1.8858;
  f[9] = 0.0013 * x ** 2 + 0.3434 * x + 2.7813;
  f[10] = 0.0026 * x ** 2 + 0.4148 * x + 3.5215;
  f[11] = 0.0038 * x ** 2 + 0.507 * x + 4.2357;
  f[12] = 0.0039 * x ** 2 + 0.72 * x + 4.2978;
  f[13] = 0.0099 * x ** 2 + 0.8173 * x + 4.5633;
  f[14] = 0.0142 * x ** 2 + 1.031 * x + 5.1436;
  f[15] = 0.0219 * x ** 2 + 1.2168 * x + 5.8647;
  f[16] = 0.1437 * x ** 2 + 2.8437 * x + 10.228;
  f[17] = 1.9048 * x ** 2 + -3.8095 * x + 44.286;

  const bands = [
    [1.2, 1.2, 1.5],
    [1.5, 1.5, 1.7],
    [1.7, 1.7, 2],
    [2, 2, 2.25],
    [2.25, 2.25, 2.5],
    [2.5, 2.5, 2.75],
    [2.75, 2.75, 3],
    [3, 3, 3.25],
    [3.25, 3.25, 3.5],
    [3.5, 3.5, 3.75],
    [3.75, 3.75, 4],
    [4, 4, 4.25],
    [4.25, 4.25, 4.5],
    [4.5, 4.5, 4.75],
    [4.75, 4.75, 7.25],
    [7.25, 7.25, 14.75],
  ];

  if (Temp < f[1]) return 1.2;
  for (let k = 2; k <= 17; k++) {
    if (Temp < f[k]) {
      const [lo, loVal, hiVal] = bands[k - 2];
      return loVal + ((Temp - f[k - 1]) / (f[k] - f[k - 1])) * (hiVal - loVal);
    }
  }
  return null; // "Out of range" in VBA
}

// ---- CVPressDrop (control valve pressure drop) -----------------------------
function ConvCVPressDrop15(TotalSucFixedPress, TotalSucVarLossRtd, TotalDischFixedPress, TotalDischVarLossRtd, OverDesign) {
  const TotalFixedPress = TotalDischFixedPress - TotalSucFixedPress;
  const TotalVarLossRtd = TotalSucVarLossRtd + TotalDischVarLossRtd;
  const TotalLossRtd = TotalFixedPress + TotalVarLossRtd;
  const PercentVarLoss1 = (TotalVarLossRtd / TotalLossRtd) * 100;
  const x = OverDesign / 100;

  let Y;
  if (PercentVarLoss1 <= 25) {
    Y = 10;
  } else if (PercentVarLoss1 <= 40) {
    const Y1 = 10;
    const Y2 = -0.0343 * x + 11.714;
    Y = Y1 + ((PercentVarLoss1 - 25) / (40 - 25)) * (Y2 - Y1);
  } else if (PercentVarLoss1 <= 50) {
    const Y1 = -0.0343 * x + 11.714;
    const Y2 = -0.0629 * x + 13.143;
    Y = Y1 + ((PercentVarLoss1 - 40) / (50 - 40)) * (Y2 - Y1);
  } else if (PercentVarLoss1 <= 75) {
    const Y1 = -0.0629 * x + 13.143;
    const Y2 = -0.108 * x + 15.429;
    Y = Y1 + ((PercentVarLoss1 - 50) / (75 - 50)) * (Y2 - Y1);
  } else if (PercentVarLoss1 <= 100) {
    const Y1 = -0.108 * x + 15.429;
    const Y2 = -0.149 * x + 17.429;
    Y = Y1 + ((PercentVarLoss1 - 75) / (100 - 75)) * (Y2 - Y1);
  } else {
    Y = -0.149 * x + 17.429;
  }
  return (Y * TotalLossRtd) / 100;
}

function Max3(a, b, c) {
  let m = a;
  if (b > m) m = b;
  if (c > m) m = c; // NOTE: VBA original uses ElseIf here (see caveat below)
  return m;
}

// NOTE ON PARITY: the VBA `Max(a,b,c)` function uses `If b > Max Then ...
// ElseIf c > Max Then ...` - i.e. c is only compared against the ORIGINAL a,
// and only when b did NOT win. Reproduced exactly (not a true max-of-3) for
// numeric parity with the desktop workbook.
function Max_VBA(a, b, c) {
  let result = a;
  if (b > result) {
    result = b;
  } else if (c > result) {
    result = c;
  }
  return result;
}

function CVPressDrop(
  Method,
  CVType,
  FlagReflux,
  TotalSucFixedPress,
  TotalSucVarLossNor,
  TotalSucVarLossRtd,
  TotalDischFixedPress,
  TotalDischVarLossNor,
  TotalDischVarLossRtd,
  OverDesign,
  M_Input
) {
  const TotalFixedPress = TotalDischFixedPress - TotalSucFixedPress;
  const TotalVarLossNor = TotalSucVarLossNor + TotalDischVarLossNor;
  const TotalVarLossRtd = TotalSucVarLossRtd + TotalDischVarLossRtd;
  const TotalLossNor = TotalFixedPress + TotalVarLossNor;
  const TotalLossRtd = TotalFixedPress + TotalVarLossRtd;
  const SucPressNor = TotalSucFixedPress - TotalSucVarLossNor;
  const SucPressRtd = TotalSucFixedPress - TotalSucVarLossRtd;

  const methodStr = String(Method);

  if (methodStr === 'No Control Valve') return 0;

  if (methodStr === 'GTP Std. Method') {
    const cvTypeStr = String(CVType);
    const BasePressDrop =
      cvTypeStr === 'single'
        ? 0.77
        : cvTypeStr === 'double'
        ? 0.49
        : cvTypeStr === 'cage'
        ? 0.28
        : cvTypeStr === 'butterfly'
        ? 0.01
        : cvTypeStr === 'vball'
        ? 0.07
        : 0.77;
    let DischPress = 1.1 * (((100 + OverDesign) / 100) ** 2 - 1) * TotalDischVarLossNor;
    DischPress = DischPress + TotalDischFixedPress + TotalDischVarLossNor;
    DischPress = (100 / 95) * (DischPress + BasePressDrop);
    return DischPress - TotalDischFixedPress - TotalDischVarLossNor;
  }

  if (methodStr === 'Lummus Method') {
    if (OverDesign <= 15) {
      const aaa = TotalVarLossNor * 0.5;
      const bbb = 1.055;
      const ccc = 0.08 * (SucPressNor + TotalLossNor);
      return Max_VBA(aaa, bbb, ccc);
    }
    const zzz = ConvCVPressDrop15(TotalSucFixedPress, TotalSucVarLossRtd, TotalDischFixedPress, TotalDischVarLossRtd, OverDesign);
    const aaa = 0.703;
    let CVPressDropRtd;
    if (String(FlagReflux) === 'Reflux') {
      CVPressDropRtd = Max_VBA(zzz, aaa, 0);
    } else {
      const bbb = 0.08 * (SucPressRtd + TotalLossRtd);
      CVPressDropRtd = Max_VBA(zzz, aaa, bbb);
    }
    return CVPressDropRtd + TotalVarLossRtd - TotalVarLossNor;
  }

  if (methodStr === 'NODCO') {
    const val1 = (TotalLossNor * 15) / 85;
    const CVPressDropRtd = val1 + TotalVarLossNor - TotalVarLossRtd;
    if (CVPressDropRtd < 0.7) {
      return 0.7 - TotalVarLossNor + TotalVarLossRtd;
    }
    return val1;
  }

  // Manual input
  return M_Input;
}



// =====================================================================
// Office.js Custom Function registrations (namespace: PUMPCALC)
// Each wrapper is a thin pass-through to the verified core function.
// =====================================================================

/**
 * Newton-Raphson-equivalent solver for RA (area ratio), used in two-phase
 * flow calcs. Fixed 1000-iteration fold, verified against VBA convergence.
 * @customfunction CALC_RA
 * @param {number} valU Unit/mode flag (1 or 2)
 * @param {number} valAS Area ratio parameter AS
 * @param {number} valAP Area ratio parameter AP
 * @returns {number} Converged RA value
 */
function CALC_RA(valU, valAS, valAP) {
  return Calc_RA(valU, valAS, valAP);
}

/**
 * Single-step J value (internal building block of the Calc_RA iteration).
 * Some sheet cells call this directly rather than going through CALC_RA.
 * @customfunction CALC_J
 * @param {number} valU Unit/mode flag (1 or 2)
 * @param {number} rAAss Assumed RA value for this iteration step
 * @param {number} valAS Area ratio parameter AS
 * @param {number} valAP Area ratio parameter AP
 * @returns {number} J value
 */
function CALC_J(valU, rAAss, valAS, valAP) {
  return Calc_J(valU, rAAss, valAS, valAP);
}

/**
 * Single-step JA (derivative) value (internal building block of the
 * Calc_RA iteration). Some sheet cells call this directly.
 * @customfunction CALC_JA
 * @param {number} valU Unit/mode flag (1 or 2)
 * @param {number} rAAss Assumed RA value for this iteration step
 * @param {number} valAS Area ratio parameter AS
 * @returns {number} JA value
 */
function CALC_JA(valU, rAAss, valAS) {
  return Calc_JA(valU, rAAss, valAS);
}

/**
 * Pipe inner diameter lookup (NPS x Schedule). Looks up a built-in
 * snapshot of the Pipe_ID sheet's data - editing the workbook's Pipe_ID
 * sheet does NOT change this function's results; new/changed pipe sizes
 * require editing pipeIdTable in this file and redeploying to GitHub Pages.
 * @customfunction PIPE_ID
 * @param {number} normID Nominal pipe size (inch)
 * @param {any} sch Schedule (e.g. "STD", "40", "80" - text or numeric, both accepted)
 * @returns {number} Inner diameter (inch)
 */
function PIPE_ID(normID, sch) {
  return Pipe_ID(normID, sch);
}

/**
 * Vapor volumetric fraction of a two-phase stream.
 * @customfunction VAPOR_VOL_FRAC
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @returns {number} Vapor volume fraction
 */
function VAPOR_VOL_FRAC(wtLiqM, wtVapM, denLiqM, denVapM) {
  return VaporVolFrac(wtLiqM, wtVapM, denLiqM, denVapM);
}

/**
 * Homogeneous-model average two-phase density.
 * @customfunction AVG_DENSITY
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @returns {number} Average density (kg/m3)
 */
function AVG_DENSITY(wtLiqM, wtVapM, denLiqM, denVapM) {
  return AvgDensity(wtLiqM, wtVapM, denLiqM, denVapM);
}

/**
 * Homogeneous-model two-phase pressure drop.
 * @customfunction DP_HOMOGENEOUS
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} e Pipe roughness (ft)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} Pressure drop (kg/cm2 per 100 ft, per original unit convention)
 */
function DP_HOMOGENEOUS(wtLiqM, wtVapM, denLiqM, denVapM, e, pipeID) {
  return DPHomogeneous(wtLiqM, wtVapM, denLiqM, denVapM, e, pipeID);
}

/**
 * Dukler-correlation two-phase pressure drop.
 * @customfunction DP_DUKLER
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} visLiq Liquid viscosity (cP)
 * @param {number} visVap Vapor viscosity (cP)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @param {number} e Pipe roughness (ft)
 * @returns {number} Pressure drop
 */
function DP_DUKLER(wtLiqM, wtVapM, denLiqM, denVapM, visLiq, visVap, pipeID, e) {
  return DPDukler(wtLiqM, wtVapM, denLiqM, denVapM, visLiq, visVap, pipeID, e);
}

/**
 * Dukler-correlation in-place (holdup) two-phase density.
 * @customfunction INPLACE_DENSITY
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} visLiq Liquid viscosity (cP)
 * @param {number} visVap Vapor viscosity (cP)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} In-place density (kg/m3)
 */
function INPLACE_DENSITY(wtLiqM, wtVapM, denLiqM, denVapM, visLiq, visVap, pipeID) {
  return InplaceDensity(wtLiqM, wtVapM, denLiqM, denVapM, visLiq, visVap, pipeID);
}

/**
 * Two-phase mixture velocity.
 * @customfunction TWO_PHASE_VELOCITY
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} Mixture velocity (m/s)
 */
function TWO_PHASE_VELOCITY(wtLiqM, wtVapM, denLiqM, denVapM, pipeID) {
  return TwoPhaseVelocity(wtLiqM, wtVapM, denLiqM, denVapM, pipeID);
}

/**
 * Baker flow-pattern map X-axis value.
 * @customfunction BAKER_XVAL
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} mWl Liquid molecular weight
 * @param {number} surTen Surface tension (0 = auto-estimate)
 * @param {number} visLiq Liquid viscosity (cP)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} Baker X value
 */
function BAKER_XVAL(wtLiqM, wtVapM, denLiqM, denVapM, mWl, surTen, visLiq, pipeID) {
  return BakerXval(wtLiqM, wtVapM, denLiqM, denVapM, mWl, surTen, visLiq, pipeID);
}

/**
 * Baker flow-pattern map Y-axis value.
 * @customfunction BAKER_YVAL
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} Baker Y value
 */
function BAKER_YVAL(wtVapM, denLiqM, denVapM, pipeID) {
  return BakerYval(wtVapM, denLiqM, denVapM, pipeID);
}

/**
 * Baker flow-pattern classification.
 * @customfunction BAKER
 * @param {number} x1 Baker X value
 * @param {number} y1 Baker Y value
 * @returns {string} Flow pattern name
 */
function BAKER(x1, y1) {
  return Baker(x1, y1);
}

/**
 * Griffith-Wallis map X-axis value (Froude number).
 * @customfunction GRIFFITH_WALLIS_XVAL
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @param {number} pipeID Pipe inner diameter (inch)
 * @returns {number} Froude number
 */
function GRIFFITH_WALLIS_XVAL(wtLiqM, wtVapM, denLiqM, denVapM, pipeID) {
  return GriffithWallisXval(wtLiqM, wtVapM, denLiqM, denVapM, pipeID);
}

/**
 * Griffith-Wallis map Y-axis value (vapor volume fraction).
 * @customfunction GRIFFITH_WALLIS_YVAL
 * @param {number} wtLiqM Liquid mass flow (kg/h)
 * @param {number} wtVapM Vapor mass flow (kg/h)
 * @param {number} denLiqM Liquid density (kg/m3)
 * @param {number} denVapM Vapor density (kg/m3)
 * @returns {number} Vapor volume fraction
 */
function GRIFFITH_WALLIS_YVAL(wtLiqM, wtVapM, denLiqM, denVapM) {
  return GriffithWallisYval(wtLiqM, wtVapM, denLiqM, denVapM);
}

/**
 * Griffith-Wallis flow-pattern classification.
 * @customfunction GRIFFITH
 * @param {number} x Froude number (from GRIFFITH_WALLIS_XVAL)
 * @param {number} y Vapor volume fraction (from GRIFFITH_WALLIS_YVAL)
 * @returns {string} Flow pattern name
 */
function GRIFFITH(x, y) {
  return Griffith(x, y);
}

/**
 * Converts head to pressure.
 * @customfunction CALC_PRES
 * @param {number} head Head (m)
 * @param {number} sg Specific gravity
 * @returns {number} Pressure (kg/cm2)
 */
function CALC_PRES(head, sg) {
  return CalcPres(head, sg);
}

/**
 * Converts pressure to head.
 * @customfunction CALC_HEAD
 * @param {number} press Pressure (kg/cm2)
 * @param {number} sg Specific gravity
 * @returns {number} Head (m)
 */
function CALC_HEAD(press, sg) {
  return CalcHead(press, sg);
}

/**
 * Estimated pump efficiency from flow rate (primary correlation).
 * @customfunction CALC_PUMP_EFF
 * @param {number} flowRate Flow rate (m3/h)
 * @returns {number} Estimated efficiency (%)
 */
function CALC_PUMP_EFF(flowRate) {
  return CalcPumpEff(flowRate);
}

/**
 * Estimated pump efficiency from flow rate, head and SG (alternate correlation).
 * Returns blank (matches VBA "Exit Function" behavior) outside its valid range
 * (Q: 25-1000 GPM).
 * @customfunction CALC_PUMP_EFF_ALT
 * @param {number} desRate Design flow rate (m3/h)
 * @param {number} spGr Specific gravity
 * @param {number} diffPress Differential pressure (kg/cm2)
 * @returns {number} Estimated efficiency (%)
 */
function CALC_PUMP_EFF_ALT(desRate, spGr, diffPress) {
  const v = CalcPumpEffAnother(desRate, spGr, diffPress);
  return v === null ? '' : v;
}

/**
 * Selects standard motor power (kW) from design duty.
 * @customfunction SEL_MOTOR_POWER
 * @param {number} desRate Design flow rate (m3/h)
 * @param {number} spGr Specific gravity
 * @param {number} diffPress Differential pressure (kg/cm2)
 * @returns {number} Selected standard motor power (kW)
 */
function SEL_MOTOR_POWER(desRate, spGr, diffPress) {
  return SelMotorPower(desRate, spGr, diffPress);
}

/**
 * Estimated NPSH required.
 * @customfunction EST_NPSHR
 * @param {number} metricQ Flow rate (m3/h)
 * @returns {number} Estimated NPSHr (m)
 */
function EST_NPSHR(metricQ) {
  return EstNPSHR(metricQ);
}

/**
 * Equivalent-length factor for fittings, from pipe ID and straight length.
 * Returns blank (matches VBA "Out of range" case) if inputs are outside the
 * calibrated table range.
 * @customfunction EQ_LEN_FACTOR
 * @param {number} pipeID Pipe inner diameter (inch)
 * @param {number} straightLength Straight pipe length (m)
 * @returns {number} Equivalent length factor
 */
function EQ_LEN_FACTOR(pipeID, straightLength) {
  const v = EqLenFactor(pipeID, straightLength);
  return v === null ? '' : v;
}

/**
 * Control valve pressure drop (method-dependent: GTP Std./Lummus/NODCO/Manual/None).
 * @customfunction CV_PRESS_DROP
 * @param {string} method Calculation method ("No Control Valve"|"GTP Std. Method"|"Lummus Method"|"NODCO"|other=Manual)
 * @param {string} cvType Valve type ("single"|"double"|"cage"|"butterfly"|"vball")
 * @param {string} flagReflux Reflux flag ("Reflux" or blank)
 * @param {number} totalSucFixedPress Suction side fixed pressure loss
 * @param {number} totalSucVarLossNor Suction side variable loss at normal flow
 * @param {number} totalSucVarLossRtd Suction side variable loss at rated flow
 * @param {number} totalDischFixedPress Discharge side fixed pressure loss
 * @param {number} totalDischVarLossNor Discharge side variable loss at normal flow
 * @param {number} totalDischVarLossRtd Discharge side variable loss at rated flow
 * @param {number} overDesign Overdesign percentage (e.g. 10 for 10%)
 * @param {number} mInput Manual input value (used only when method is not recognized)
 * @returns {number} Control valve pressure drop
 */
function CV_PRESS_DROP(
  method,
  cvType,
  flagReflux,
  totalSucFixedPress,
  totalSucVarLossNor,
  totalSucVarLossRtd,
  totalDischFixedPress,
  totalDischVarLossNor,
  totalDischVarLossRtd,
  overDesign,
  mInput
) {
  return CVPressDrop(
    method,
    cvType,
    flagReflux,
    totalSucFixedPress,
    totalSucVarLossNor,
    totalSucVarLossRtd,
    totalDischFixedPress,
    totalDischVarLossNor,
    totalDischVarLossRtd,
    overDesign,
    mInput
  );
}

// Explicit registration (namespace applied via manifest.xml, not here)
if (typeof CustomFunctions !== 'undefined') {
  CustomFunctions.associate('CALC_RA', CALC_RA);
  CustomFunctions.associate('CALC_J', CALC_J);
  CustomFunctions.associate('CALC_JA', CALC_JA);
  CustomFunctions.associate('PIPE_ID', PIPE_ID);
  CustomFunctions.associate('VAPOR_VOL_FRAC', VAPOR_VOL_FRAC);
  CustomFunctions.associate('AVG_DENSITY', AVG_DENSITY);
  CustomFunctions.associate('DP_HOMOGENEOUS', DP_HOMOGENEOUS);
  CustomFunctions.associate('DP_DUKLER', DP_DUKLER);
  CustomFunctions.associate('INPLACE_DENSITY', INPLACE_DENSITY);
  CustomFunctions.associate('TWO_PHASE_VELOCITY', TWO_PHASE_VELOCITY);
  CustomFunctions.associate('BAKER_XVAL', BAKER_XVAL);
  CustomFunctions.associate('BAKER_YVAL', BAKER_YVAL);
  CustomFunctions.associate('BAKER', BAKER);
  CustomFunctions.associate('GRIFFITH_WALLIS_XVAL', GRIFFITH_WALLIS_XVAL);
  CustomFunctions.associate('GRIFFITH_WALLIS_YVAL', GRIFFITH_WALLIS_YVAL);
  CustomFunctions.associate('GRIFFITH', GRIFFITH);
  CustomFunctions.associate('CALC_PRES', CALC_PRES);
  CustomFunctions.associate('CALC_HEAD', CALC_HEAD);
  CustomFunctions.associate('CALC_PUMP_EFF', CALC_PUMP_EFF);
  CustomFunctions.associate('CALC_PUMP_EFF_ALT', CALC_PUMP_EFF_ALT);
  CustomFunctions.associate('SEL_MOTOR_POWER', SEL_MOTOR_POWER);
  CustomFunctions.associate('EST_NPSHR', EST_NPSHR);
  CustomFunctions.associate('EQ_LEN_FACTOR', EQ_LEN_FACTOR);
  CustomFunctions.associate('CV_PRESS_DROP', CV_PRESS_DROP);
}

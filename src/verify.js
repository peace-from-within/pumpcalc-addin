const fs = require('fs');
const core = require('./core.js');

const cases = JSON.parse(fs.readFileSync('./verify_cases.json', 'utf-8'));

const fnMap = {
  Calc_RA: core.Calc_RA,
  VaporVolFrac: core.VaporVolFrac,
  AvgDensity: core.AvgDensity,
  DPHomogeneous: core.DPHomogeneous,
  DPDukler: core.DPDukler,
  InplaceDensity: core.InplaceDensity,
  TwoPhaseVelocity: core.TwoPhaseVelocity,
  BakerXval: core.BakerXval,
  BakerYval: core.BakerYval,
  Baker: core.Baker,
  GriffithWallisXval: core.GriffithWallisXval,
  GriffithWallisYval: core.GriffithWallisYval,
  Griffith: core.Griffith,
  CalcPres: core.CalcPres,
  CalcHead: core.CalcHead,
  CalcPumpEff: core.CalcPumpEff,
  EstNPSHR: core.EstNPSHR,
  CVPressDrop: core.CVPressDrop,
};

let totalTests = 0, totalPass = 0;
const failures = [];

for (const [fnName, tests] of Object.entries(cases)) {
  const fn = fnMap[fnName];
  if (!fn) { console.log('NO JS FN FOR', fnName); continue; }
  let pass = 0;
  for (const t of tests) {
    totalTests++;
    let actual;
    try {
      actual = fn(...t.args);
    } catch (e) {
      failures.push({ fn: fnName, args: t.args, expect: t.expect, error: String(e) });
      continue;
    }
    const expect = t.expect;
    let ok;
    if (typeof expect === 'string' || typeof actual === 'string') {
      ok = String(actual) === String(expect);
    } else {
      const denom = Math.max(Math.abs(expect), 1e-9);
      ok = Math.abs(actual - expect) / denom < 1e-6 || Math.abs(actual - expect) < 1e-6;
    }
    if (ok) { pass++; totalPass++; }
    else failures.push({ fn: fnName, args: t.args, expect, actual });
  }
  console.log(fnName, `${pass}/${tests.length}`);
}

console.log('---');
console.log('TOTAL', `${totalPass}/${totalTests}`);
if (failures.length) {
  console.log('FAILURES (first 20):');
  for (const f of failures.slice(0, 20)) {
    console.log(JSON.stringify(f));
  }
}

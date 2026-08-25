const fs = require('fs');
const core = require('./core.js');
const cases = JSON.parse(fs.readFileSync('./verify_cases2.json', 'utf-8'));

const fnMap = {
  CalcPumpEffAnother: core.CalcPumpEffAnother,
  SelMotorPower: core.SelMotorPower,
  EqLenFactor: core.EqLenFactor,
};

let totalTests = 0, totalPass = 0;
const failures = [];
for (const [fnName, tests] of Object.entries(cases)) {
  const fn = fnMap[fnName];
  let pass = 0;
  for (const t of tests) {
    totalTests++;
    const actual = fn(...t.args);
    const expect = t.expect;
    let ok;
    if (expect === null || actual === null || actual === undefined) {
      ok = (expect === null) === (actual === null || actual === undefined);
    } else {
      const denom = Math.max(Math.abs(expect), 1e-9);
      ok = Math.abs(actual - expect) / denom < 1e-6 || Math.abs(actual - expect) < 1e-6;
    }
    if (ok) { pass++; totalPass++; }
    else failures.push({ fn: fnName, args: t.args, expect, actual });
  }
  console.log(fnName, `${pass}/${tests.length}`);
}
console.log('TOTAL', `${totalPass}/${totalTests}`);
if (failures.length) {
  console.log('FAILURES:');
  failures.slice(0, 15).forEach(f => console.log(JSON.stringify(f)));
}

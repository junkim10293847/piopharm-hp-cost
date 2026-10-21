#!/usr/bin/env node
// 화성·평택 매출원가 명령줄 도구
//   node cli/hp-cost.js <기초파일.xlsx | 폴더> [-o 출력경로] [--month 9] [--fx 1360] [--ratio ratio.json] [--json]
// 폴더를 주면 그 안의 .xlsx를 모두 처리해 <폴더>/output/ 에 결과를 저장합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const Engine = require('../src/engine.js');
const bookFromWorkbook = require('../src/parse.js');
const Report = require('../src/report.js');

function parseArgs(argv) {
  const a = { inputs: [] };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '-o' || k === '--out') a.out = argv[++i];
    else if (k === '--month') a.month = Number(argv[++i]);
    else if (k === '--fx') a.fx = Number(argv[++i]);
    else if (k === '--ratio') a.ratio = JSON.parse(fs.readFileSync(argv[++i], 'utf8'));
    else if (k === '--json') a.json = true;
    else if (k === '-h' || k === '--help') a.help = true;
    else a.inputs.push(k);
  }
  return a;
}

function processFile(file, args, outPath) {
  const wb = XLSX.readFile(file, { cellFormula: true, cellDates: true });
  const R = Engine.run(bookFromWorkbook(XLSX, wb), { month: args.month, fx: args.fx, ratio: args.ratio });
  const out = outPath || path.join(path.dirname(file), Report.outputName(R));
  XLSX.writeFile(Report.buildWorkbook(XLSX, R, path.basename(file)), out);
  const S = Report.totals(R);
  const summary = {
    file: path.basename(file), output: out, month: R.opts.month, fx: R.opts.fx,
    total: Math.round(R.totalShip),
    pyeongtaek: { sales: Math.round(S.P.매출), raw: Math.round(S.P.원료), material: Math.round(S.P.자재) },
    hwaseong: { sales: Math.round(S.H.매출), raw: Math.round(S.H.원료), material: Math.round(S.H.자재) },
    lots: R.lots.length, items: R.hw.length, flags: Report.flagsOf(R).length,
  };
  return summary;
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.inputs.length) {
    console.log('사용법: node cli/hp-cost.js <기초파일.xlsx | 폴더> [-o 출력] [--month 9] [--fx 1360] [--ratio ratio.json] [--json]');
    process.exit(args.help ? 0 : 1);
  }
  const results = []; let failed = 0;
  for (const input of args.inputs) {
    const isDir = fs.statSync(input).isDirectory();
    const files = isDir
      ? fs.readdirSync(input).filter(f => /\.xlsx?$/i.test(f) && !f.startsWith('~$') && !f.startsWith('화성평택_매출원가_')).map(f => path.join(input, f))
      : [input];
    const outDir = isDir ? (args.out || path.join(input, 'output')) : null;
    if (outDir) fs.mkdirSync(outDir, { recursive: true });
    for (const f of files) {
      try {
        const out = outDir ? path.join(outDir, path.basename(f).replace(/\.xlsx?$/i, '') + '_결과.xlsx') : args.out;
        const s = processFile(f, args, out);
        results.push(s);
        if (!args.json) {
          const w = n => n.toLocaleString('ko-KR');
          console.log(`✔ ${s.file} (${s.month}월) → ${s.output}`);
          console.log(`  총 매출 ${w(s.total)} | 평택 ${w(s.pyeongtaek.sales)} | 화성 ${w(s.hwaseong.sales)} | 확인사항 ${s.flags}건`);
        }
      } catch (e) {
        failed++;
        console.error(`✘ ${path.basename(f)}: ${e.message}`);
      }
    }
  }
  if (args.json) console.log(JSON.stringify(results, null, 2));
  process.exit(failed ? 1 : 0);
}

if (require.main === module) main();
module.exports = { processFile };

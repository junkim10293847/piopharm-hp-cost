#!/usr/bin/env node
// 화성·평택 매출원가 명령줄 도구
//   node cli/hp-cost.js <기초파일.xlsx | 폴더 | 개별 파일들…> [-o 출력경로] [--month 9] [--ratio ratio.json] [--json]
// 폴더를 주면 그 안의 .xlsx를 모두 처리해 <폴더>/output/ 에 결과를 저장합니다.
// 혼자서 시트가 다 갖춰진 파일은 따로 계산하고, 시트별로 따로 받은 파일들은 하나로 묶어 계산합니다.
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const Engine = require('../src/engine.js');
const { bookFromWorkbooks } = require('../src/parse.js');
const Report = require('../src/report.js');

function parseArgs(argv) {
  const a = { inputs: [] };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '-o' || k === '--out') a.out = argv[++i];
    else if (k === '--month') a.month = Number(argv[++i]);
    else if (k === '--ratio') a.ratio = JSON.parse(fs.readFileSync(argv[++i], 'utf8'));
    else if (k === '--json') a.json = true;
    else if (k === '-h' || k === '--help') a.help = true;
    else a.inputs.push(k);
  }
  return a;
}

const readBook = file => XLSX.readFile(file, { cellFormula: true, cellDates: true });

// files: 하나의 기초 자료로 묶을 파일들
function processFiles(files, args, outPath) {
  const B = bookFromWorkbooks(XLSX, files.map(f => ({ name: path.basename(f), wb: readBook(f) })));
  if (B.missing.length) throw new Error('빠진 자료: ' + B.missing.join(', '));
  const R = Engine.run(B, { month: args.month, ratio: args.ratio });
  const label = files.map(f => path.basename(f)).join(', ');
  const out = outPath || path.join(path.dirname(files[0]), Report.outputName(R));
  XLSX.writeFile(Report.buildWorkbook(XLSX, R, label), out);
  const S = Report.totals(R);
  return {
    file: label, output: out, month: R.opts.month,
    total: Math.round(R.totalShip),
    pyeongtaek: { sales: Math.round(S.P.매출), raw: Math.round(S.P.원료), material: Math.round(S.P.자재) },
    hwaseong: { sales: Math.round(S.H.매출), raw: Math.round(S.H.원료), material: Math.round(S.H.자재) },
    pyeongtaekPrevMonth: Math.round(R.P전월.매출),
    pyeongtaekStock: R.stock.hasFco ? Math.round(R.stock.rows.filter(r => r.공장 === '평택').reduce((a, r) => a + r.현재고금액, 0)) : null,
    lots: R.lots.length, prevLots: R.prev.length, items: R.hw.length, flags: Report.flagsOf(R).length,
  };
}
const processFile = (file, args, outPath) => processFiles([file], args, outPath);

// 혼자 완결된 파일은 각각, 나머지(시트별 개별 파일)는 한 묶음으로
function groupFiles(files) {
  const solo = [], parts = [];
  for (const f of files) {
    let ok = false;
    try { ok = !bookFromWorkbooks(XLSX, [{ name: path.basename(f), wb: readBook(f) }]).missing.length; } catch (e) { /* 아래에서 오류 표시 */ }
    (ok ? solo : parts).push(f);
  }
  return [...solo.map(f => [f]), ...(parts.length ? [parts] : [])];
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.inputs.length) {
    console.log('사용법: node cli/hp-cost.js <기초파일.xlsx | 폴더 | 개별 파일들…> [-o 출력] [--month 9] [--ratio ratio.json] [--json]');
    process.exit(args.help ? 0 : 1);
  }
  const results = []; let failed = 0;
  const gone = args.inputs.filter(i => !fs.existsSync(i));
  if (gone.length) { console.error('✘ 찾을 수 없는 경로: ' + gone.join(', ')); process.exit(1); }
  const dirs = args.inputs.filter(i => fs.statSync(i).isDirectory());
  const loose = args.inputs.filter(i => !dirs.includes(i));
  const xl = d => fs.readdirSync(d).filter(f => /\.xlsx?$/i.test(f) && !f.startsWith('~$') && !f.startsWith('화성평택_매출원가_')).map(f => path.join(d, f));
  const jobs = [
    ...dirs.map(d => ({ files: xl(d), outDir: args.out || path.join(d, 'output') })),
    ...(loose.length ? [{ files: loose, outDir: null }] : []),
  ];
  for (const job of jobs) {
    if (job.outDir) fs.mkdirSync(job.outDir, { recursive: true });
    for (const group of groupFiles(job.files)) {
      try {
        let out = args.out && !job.outDir ? args.out : null;
        if (job.outDir) out = group.length === 1
          ? path.join(job.outDir, path.basename(group[0]).replace(/\.xlsx?$/i, '') + '_결과.xlsx')
          : null;
        const s = processFiles(group, args, out);
        if (job.outDir && group.length > 1) {
          const dest = path.join(job.outDir, path.basename(s.output));
          fs.renameSync(s.output, dest); s.output = dest;
        }
        results.push(s);
        if (!args.json) {
          const w = n => n.toLocaleString('ko-KR');
          console.log(`✔ ${group.length > 1 ? `개별 파일 ${group.length}개` : s.file} (${s.month}월) → ${s.output}`);
          console.log(`  총 매출 ${w(s.total)} | 평택 ${w(s.pyeongtaek.sales)} (전월 생산 ${w(s.pyeongtaekPrevMonth)}) | 화성 ${w(s.hwaseong.sales)} | 평택 미출고 재고 ${s.pyeongtaekStock == null ? "FCO 없음" : w(s.pyeongtaekStock)} | 확인사항 ${s.flags}건`);
        }
      } catch (e) {
        failed++;
        console.error(`✘ ${group.map(f => path.basename(f)).join(', ')}: ${e.message}`);
      }
    }
  }
  if (args.json) console.log(JSON.stringify(results, null, 2));
  process.exit(failed ? 1 : 0);
}

if (require.main === module) main();
module.exports = { processFile, processFiles, groupFiles };

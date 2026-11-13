const test = require('node:test');
const assert = require('node:assert/strict');
const XLSX = require('xlsx');
const Engine = require('../src/engine.js');
const bookFromWorkbook = require('../src/parse.js');
const Report = require('../src/report.js');
const { fixtureWorkbook } = require('../fixtures/sample.js');

const R = Engine.run(bookFromWorkbook(XLSX, fixtureWorkbook()));
const lot = id => R.lots.find(o => o.로트번호 === id);
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 1e-6, `${msg}: ${a} ≠ ${b}`);

test('로트목록: 당월 4개 + 전월 구간 1개, 당월 자동 인식, 생산구분 → 평택비율', () => {
  assert.equal(R.lots.length, 5);
  assert.deepEqual(R.lots.map(o => o.구분), ['당월', '당월', '당월', '당월', '전월']);
  assert.equal(R.opts.month, 9);
  assert.equal(lot('T001').K, 1); assert.equal(lot('T003').K, 0.8);
});

test('생산공장·포장공장 열로 평택비율 (미가공 로트목록 형식)', () => {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
    ['No.', '제품코드\n(제공자료)', '로트번호\n(제공자료)', '제품명', '고객사', '규격', '단가', '완료수량', '완료금액', '비고', '포장유형', '제형', '생산공장', '포장공장'],
    [1, '1AA0000001', 'A1', '가', '홍길동', 30, 100, 10, 1000, '', '', '', '평택', '평택'],
    [2, '1AA0000002', 'A2', '나', '홍길동', 30, 100, 10, 1000, '', '', '', '평택', '화성'],
  ]), '9월 ');
  const L = Engine.readLots(bookFromWorkbook(XLSX, wb)).lots;
  assert.deepEqual(L.map(l => l.생산구분), ['평택', '화성포장']);
  assert.equal(L[0].완료단가, 100);
});

test('같은 제품 여러 로트: 생산수량 비율로 원가, 출하분만', () => {
  near(lot('T001').개당원가, 250, '1개당 원가 = 1,000,000 ÷ 4,000');
  near(lot('T001').원가, 250000, 'T001 원가'); near(lot('T001').원료, 175000, 'T001 원료');
  near(lot('T002').자재, 225000, 'T002 자재');
});

test('화성포장 로트: 평택 80% / 화성 20%', () => {
  const o = lot('T003');
  near(o.매출, 800000, '평택분 매출'); near(o.원료, 320000, '평택분 원료'); near(o.자재, 80000, '평택분 자재');
  near(o.H매출, 200000, '화성분 매출'); near(o.H원료, 80000, '화성분 원료'); near(o.H자재, 20000, '화성분 자재');
});

test('당월 단가대장에 없는 로트: 이전 달 대장 kg당 원가 × 출하중량', () => {
  const o = lot('T004');
  assert.equal(o.원가월, 8);
  near(o.원가, 1400000, '140kg × 10,000'); near(o.원료, 1000000, '원료'); near(o.자재, 400000, '자재');
});

test('출하 안 된 로트: 매출·원가 0, 잔량을 재료비로 남김', () => {
  const o = lot('T006');
  near(o.매출, 0, '매출'); near(o.원가, 0, '원가'); near(o.잔량, 2000, '잔량'); near(o.잔량재료비, 300000, '잔량 재료비');
  assert.ok(o.flags.some(f => f.startsWith('출하 없음')));
});

test('로트목록에 없는 평택 출하(작업지시 10-30) → 전월 생산, 생산월은 보관함 이전 달 로트목록으로만', () => {
  assert.equal(R.prev.length, 1);
  const o = R.prev[0];
  assert.equal(o.로트번호, 'T005'); assert.equal(o.원가, null, '이전 달 로트목록이 없으면 빈칸');
  assert.ok(o.flags.some(f => f.includes('생산월 확인 불가')));
  const R2 = Engine.run(bookFromWorkbook(XLSX, fixtureWorkbook()), { history: { lotMonth: new Map([['1AA0000003|T005', 8]]) } });
  const p = R2.prev[0];
  assert.equal(p.원가월, 8); assert.match(p.생산월근거, /보관함/);
  near(p.매출, 390000, '매출'); near(p.원가, 560000, '56kg × 10,000'); near(p.원료, 400000, '원료');
});

test('추정하지 않음: 생산월 단가대장이 없으면 원가 빈칸(당월 대장으로 대신하지 않음)', () => {
  const wb = fixtureWorkbook(); wb.SheetNames = wb.SheetNames.filter(n => !n.startsWith('8월'));
  const R2 = Engine.run(bookFromWorkbook(XLSX, wb));
  const t4 = R2.lots.find(o => o.로트번호 === 'T004'), t5 = R2.prev.find(o => o.로트번호 === 'T005');
  assert.equal(t4.원가, null); assert.equal(t5.원가, null);
  assert.ok(t4.flags.some(f => f.includes('원가 빈칸'))); assert.ok(t5.flags.some(f => f.includes('원가 빈칸')));
  near(R2.미확인.평택, 975000 + 390000, '원가 미확인 평택 매출');
  assert.ok(t4.flags.some(f => f.includes('생산월 8월')));
});

test('화성 품목: 출하중량 × kg당 제조원가', () => {
  const g = R.hw.find(x => x.품목코드 === '1BB0000001');
  near(g.매출, 7500000, '매출'); near(g.중량, 150, '중량'); near(g.kg당원가, 10000, 'kg당 원가');
  near(g.원가, 1500000, '원가'); near(g.원료, 1125000, '원료'); near(g.자재, 375000, '자재');
  const x = R.hw.find(x => x.품목코드 === '4C00009999');
  assert.equal(x.원가, null); assert.match(x.flag, /제품 외/);
});

test('평택 + 화성 = 출하대장 합계', () => {
  near(R.totalShip, 14165000, '총 출하');
  near(R.P당월.매출, 5775000, '평택 당월 매출'); near(R.P당월.원료, 2020000, '평택 당월 원료'); near(R.P당월.자재, 780000, '평택 당월 자재');
  near(R.P.매출, 6165000, '평택 매출(전월 포함)');
  near(R.H.매출, 8000000, '화성 매출');
  near(R.P.매출 + R.H.매출, R.totalShip, '합계 일치');
  near(R.Hu.미확인매출, 300000, '원가 미확인');
});

test('미출고 재고: FCO 현재고 > 0 제품만, 작업지시로 공장 구분', () => {
  const rows = R.stock.rows;
  assert.equal(R.stock.hasFco, true);
  assert.deepEqual(rows.map(r => r.제품코드).sort(), ['1AA0000003', '1AA0000004', '1BB0000001']);
  const pt = rows.filter(r => r.공장 === '평택');
  near(pt.reduce((a, r) => a + r.현재고금액, 0), 1350000, '평택 기말 제품재고');
  assert.equal(rows.find(r => r.제품코드 === '1BB0000001').공장, '화성');
  near(rows.find(r => r.제품코드 === '1AA0000004').재료비추정, 300000, '재료비 추정');
  assert.equal(rows.find(r => r.제품코드 === '1AA0000004').구분, '당월 생산·미출고');
});

test('FCO가 없으면 로트 잔량으로 추정', () => {
  const wb = fixtureWorkbook(); wb.SheetNames = wb.SheetNames.filter(n => !n.includes('FCO'));
  const R2 = Engine.run(bookFromWorkbook(XLSX, wb));
  assert.equal(R2.stock.hasFco, false);
  assert.equal(R2.stock.rows.length, 1); near(R2.stock.rows[0].재료비추정, 300000, '추정');
});

test('엑셀 보고서: 시트 6개, 요약 합계 행, 버전 정보', () => {
  const wb = Report.buildWorkbook(XLSX, R, 'fixture.xlsx');
  assert.deepEqual(wb.SheetNames, ['요약', '평택_당월로트', '평택_전월생산', '화성_품목별', '미출고_재고', '확인사항']);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets['요약'], { header: 1 });
  assert.equal(rows[7][0], '합계'); assert.equal(rows[7][1], 14165000);
  assert.ok(rows.some(r => r[0] === '계산기 버전'));
  assert.equal(XLSX.utils.sheet_to_json(wb.Sheets['평택_당월로트']).length, 5);
});

test('사용자 설정: 비율 직접 지정', () => {
  const R2 = Engine.run(bookFromWorkbook(XLSX, fixtureWorkbook()), { ratio: { '화성포장': 0.5 } });
  near(R2.lots.find(o => o.로트번호 === 'T003').매출, 500000, '비율 지정');
});

// ---------- 시트별 개별 파일 ----------
const { bookFromWorkbooks } = require('../src/parse.js');
// 통합 기초 파일을 시트마다 따로 저장한 것처럼 나눔 (ERP에서 하나씩 내려받은 상태)
function splitFixture(nameOf) {
  const wb = fixtureWorkbook();
  return wb.SheetNames.map(sn => {
    const one = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(one, wb.Sheets[sn], 'Sheet1');
    const buf = XLSX.write(one, { type: 'buffer', bookType: 'xlsx' });
    return { name: nameOf(sn), wb: XLSX.read(buf, { type: 'buffer', cellFormula: true, cellDates: true }) };
  });
}
const sameResult = (R2) => {
  const a = Report.totals(R), b = Report.totals(R2);
  near(b.P.매출, a.P.매출, '평택 매출'); near(b.H.매출, a.H.매출, '화성 매출');
  near(b.P.원료, a.P.원료, '평택 원료'); near(b.H.자재, a.H.자재, '화성 자재');
  assert.equal(R2.lots.length, R.lots.length);
  assert.equal(R2.opts.month, R.opts.month);
};

test('개별 파일: 시트 이름이 Sheet1이어도 파일 이름으로 인식, 통합 파일과 결과 동일', () => {
  const files = splitFixture(sn => sn + '.xlsx');
  const B = bookFromWorkbooks(XLSX, files);
  assert.deepEqual(B.missing, []);
  sameResult(Engine.run(B));
});

test('개별 파일: 파일 이름에 단서가 없어도 머리글로 인식(월은 파일 이름에서)', () => {
  let i = 0;
  const files = splitFixture(sn => (sn.match(/\d+월/) ? sn.match(/\d+월/)[0] + '_' : '') + 'ERP_export_' + (i++) + '.xlsx');
  const B = bookFromWorkbooks(XLSX, files);
  assert.deepEqual(B.missing, []);
  assert.ok(B.recognized.every(r => r.how === '머리글'));
  sameResult(Engine.run(B));
});

test('개별 파일: 빠진 파일 알려줌, 같은 종류 중복은 첫 번째만 사용', () => {
  const files = splitFixture(sn => sn + '.xlsx').filter(f => !f.name.includes('발주'));
  const B = bookFromWorkbooks(XLSX, [...files, files.find(f => f.name.includes('출하대장'))]);
  assert.deepEqual(B.missing, ['발주 현황']);
  assert.equal(B.recognized.filter(r => r.dup).length, 1);
});

test('개별 파일: 통합 파일 1개를 넣어도 그대로 동작', () => {
  const B = bookFromWorkbooks(XLSX, [{ name: '기초.xlsx', wb: fixtureWorkbook() }]);
  assert.deepEqual(B.missing, []);
  sameResult(Engine.run(B));
});

// ---------- 자료별 칸 ----------
const { SLOTS, pickSheet, bookFromSlots } = require('../src/parse.js');
const slotOfSheet = sn => /발주/.test(sn) ? 'order' : /작업지시/.test(sn) ? 'wo' : /출하대장/.test(sn) ? 'ship'
  : /9월 제품 단가/.test(sn) ? 'ledCur' : /8월 제품 단가/.test(sn) ? 'ledPrev' : /FCO/.test(sn) ? 'fco' : /당월 생산/.test(sn) ? 'lots' : null;

test('칸: 파일·시트 이름과 상관없이 칸에 넣은 대로 계산, 통합 파일과 결과 동일', () => {
  const filled = {};
  splitFixture(() => '아무이름.xlsx').forEach((f, i) => {
    const id = slotOfSheet(fixtureWorkbook().SheetNames[i]);
    if (!id) return;
    const slot = SLOTS.find(s => s.id === id);
    filled[id] = { wb: f.wb, file: 'export' + i + '.xlsx', sheet: pickSheet(XLSX, f.wb, slot, 9, false) };
  });
  const B = bookFromSlots(XLSX, filled, 9);
  assert.deepEqual(B.missing, []); assert.deepEqual(B.warn, []);
  sameResult(Engine.run(B, { month: 9 }));
});

test('칸: 통합 파일 1개를 넣으면 맞는 칸에 자동 배치', () => {
  const wb = fixtureWorkbook(), filled = {};
  for (const slot of SLOTS) { const sh = pickSheet(XLSX, wb, slot, 9, true); if (sh) filled[slot.id] = { wb, sheet: sh, file: '기초.xlsx' }; }
  assert.equal(filled.ledCur.sheet, '9월 제품 단가 대장');
  assert.equal(filled.ledPrev.sheet, '8월 제품 단가 대장');
  assert.deepEqual(Object.keys(filled).sort(), ['fco', 'ledCur', 'ledPrev', 'lots', 'order', 'ship', 'wo']);
  sameResult(Engine.run(bookFromSlots(XLSX, filled, 9), { month: 9 }));
});

test('칸: 잘못된 칸에 넣으면 빠진 열을 알려줌, 전월 단가 대장은 선택', () => {
  const wb = fixtureWorkbook();
  const B = bookFromSlots(XLSX, { order: { wb, file: 'x.xlsx', sheet: '자재투입현황(6-9)'.replace(/\(.*/, '') } }, 9);
  assert.ok(!B.missing.includes('FCO 수불'));
  assert.ok(B.missing.includes('출하대장'));
  assert.ok(!B.missing.some(m => m.includes('전월')));
  const sn = wb.SheetNames.find(n => n.includes('자재투입'));
  const B2 = bookFromSlots(XLSX, { order: { wb, file: 'x.xlsx', sheet: sn } }, 9);
  assert.deepEqual(B2.warn[0].missing, ['LOT', '매출월', '수주단가']);
});

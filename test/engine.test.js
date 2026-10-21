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

test('로트목록: 당월 3개 + 전월 1개, 당월·환율 자동 인식', () => {
  assert.equal(R.lots.length, 4);
  assert.deepEqual(R.lots.map(o => o.구분), ['당월', '당월', '당월', '전월']);
  assert.equal(R.opts.month, 9);
  assert.equal(R.opts.fx, 1300);
  assert.equal(lot('T004').환율적용, true);
  assert.equal(lot('T001').환율적용, false);
  near(lot('T004').보정계수, 0.5, '보정계수(AX열)');
});

test('발주현황: 당월(매출 월=9)만 집계', () => {
  assert.equal(lot('T001').OQty, 1000);
  near(lot('T004').OPrice, 1.5 * 1300, '외화 단가 × 환율');
});

test('작업지시: 제품코드+로트 복합키로 규격을 찾음(반제품 행 무시)', () => {
  assert.equal(lot('T001').Spec, '1,000mg x 30연질캡슐/set');
  near(lot('T001').AS, 1000, '생산중량→set');
});

test('규격 문자열 해석', () => {
  assert.deepEqual(Engine.parseSpec('1,000mg x 30연질캡슐 x 3ea/set'), { mg: 1000, cap: 30, ea: 3 });
  assert.deepEqual(Engine.parseSpec('20g x 14포/set'), { mg: 20000, cap: 14, ea: 1 });
  assert.deepEqual(Engine.parseSpec('15ml x 10포/set'), { mg: 15000, cap: 10, ea: 1 });
  assert.deepEqual(Engine.parseSpec('1,256mg/연질캡슐'), { mg: 1256, cap: 1, ea: 1 });
  assert.deepEqual(Engine.parseSpec('660mg x 120정 x 2병/set'), { mg: 660, cap: 120, ea: 2 });
});

test('같은 제품 여러 로트: 완료수량 비율로 원가 배분', () => {
  near(lot('T001').Share, 0.25, 'T001 배분'); near(lot('T002').Share, 0.75, 'T002 배분');
  near(lot('T001').AA, 250000, 'T001 제품금액'); near(lot('T001').AD, 175000, 'T001 원료');
  near(lot('T002').AE, 225000, 'T002 자재');
});

test('화성포장 로트: 평택 80% / 화성 20%', () => {
  const o = lot('T003');
  near(o.X, 800000, '평택분 매출'); near(o.AD, 320000, '평택분 원료'); near(o.AE, 80000, '평택분 자재');
  near(o.HX, 200000, '화성분 매출'); near(o.HAD, 80000, '화성분 원료'); near(o.HAE, 20000, '화성분 자재');
  assert.ok(o.flags.includes('생산-수주 수량차이'));
});

test('전월 로트: 제조월(8월) 단가대장 + 보정계수를 원가 전체에 적용', () => {
  const o = lot('T004');
  assert.equal(o.CM, 8);
  near(o.AA, 1400000, '제품금액'); near(o.AD, 1000000, '원료'); near(o.AE, 400000, '자재');
  near(o.AJ, 140, '생산중량 × 보정'); near(o.AW, 1, '조정비율');
});

test('수율: 같은 투입번호는 한 번만 합산', () => {
  near(lot('T001').AL, 30 / 31, 'T001 수율');
});

test('화성 품목: 출하중량 × kg당 제조원가', () => {
  const g = R.hw.find(x => x.품목코드 === '1BB0000001');
  near(g.매출, 7500000, '매출'); near(g.중량, 150, '중량'); near(g.kg당원가, 10000, 'kg당 원가');
  near(g.원가, 1500000, '원가'); near(g.원료, 1125000, '원료'); near(g.자재, 375000, '자재');
  const x = R.hw.find(x => x.품목코드 === '4C00009999');
  assert.equal(x.원가, null); assert.match(x.flag, /제품 외/);
});

test('평택 + 화성 = 출하대장 합계', () => {
  near(R.totalShip, 13775000, '총 출하');
  near(R.P.매출, 5775000, '평택 매출'); near(R.P.원료, 2020000, '평택 원료'); near(R.P.자재, 780000, '평택 자재');
  near(R.H.매출, 8000000, '화성 매출');
  near(R.P.매출 + R.H.매출, R.totalShip, '합계 일치');
  near(R.Hu.미확인매출, 300000, '원가 미확인');
});

test('엑셀 보고서: 시트 4개, 요약 합계 행', () => {
  const wb = Report.buildWorkbook(XLSX, R, 'fixture.xlsx');
  assert.deepEqual(wb.SheetNames, ['요약', '평택_로트별', '화성_품목별', '확인사항']);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets['요약'], { header: 1 });
  assert.equal(rows[5][0], '합계'); assert.equal(rows[5][1], 13775000);
  const pt = XLSX.utils.sheet_to_json(wb.Sheets['평택_로트별']);
  assert.equal(pt.length, 4);
});

test('사용자 설정: 당월·환율·비율 직접 지정', () => {
  const R2 = Engine.run(bookFromWorkbook(XLSX, fixtureWorkbook()), { fx: 1400, ratio: { '평택': 1, '화성포장': 0.5 } });
  near(R2.lots.find(o => o.로트번호 === 'T004').OPrice, 1.5 * 1400, '환율 지정');
  near(R2.lots.find(o => o.로트번호 === 'T003').X, 500000, '비율 지정');
});

// SheetJS 통합문서 → 엔진 입력 (A1 기준 2차원 배열 + 수식 조회)
const _sheetCache = typeof WeakMap !== 'undefined' ? new WeakMap() : null;
function sheetsOf(XLSX, wb) {
  if (_sheetCache && _sheetCache.has(wb)) return _sheetCache.get(wb);
  const out = _sheetsOf(XLSX, wb);
  if (_sheetCache) _sheetCache.set(wb, out);
  return out;
}
function _sheetsOf(XLSX, wb) {
  return wb.SheetNames.map(name => {
    const ws = wb.Sheets[name];
    const rows = [];
    if (ws && ws['!ref']) {
      const rg = XLSX.utils.decode_range(ws['!ref']);
      for (let R = 0; R <= rg.e.r; R++) {
        const row = new Array(rg.e.c + 1).fill(null);
        for (let C = rg.s.c; C <= rg.e.c; C++) {
          const c = ws[XLSX.utils.encode_cell({ r: R, c: C })];
          if (c && c.t !== 'e' && c.t !== 'z') row[C] = c.v;
        }
        rows.push(row);
      }
    }
    return { name, rows, formula: (r, c) => { const x = ws[XLSX.utils.encode_cell({ r, c })]; return x && x.f ? String(x.f) : null; } };
  });
}

function bookFromWorkbook(XLSX, wb) {
  return { sheets: sheetsOf(XLSX, wb) };
}

// ---------- 여러 파일을 하나의 기초 자료로 ----------
// 시트 종류는 ① 시트 이름 → ② 머리글 → ③ 파일 이름 순서로 알아냅니다.
// 시트 이름에 종류가 없으면 엔진이 찾을 수 있게 '9월 출하대장'처럼 이름을 붙여 줍니다.
const SHEET_KINDS = [
  { kw: '발주현황', label: '발주 현황', need: ['LOT', '매출월', '수주단가'] },
  { kw: '작업지시', label: '작업지시현황', need: ['제품로트번호', '지시제품'] },
  { kw: '출하대장', label: '출하대장', need: ['로트번호', '매출업체', '공급가액'], monthly: true },
  { kw: '단가대장', label: '제품 단가 대장', need: ['자재코드', '상세금액'], monthly: true, many: true },
  { kw: '자재투입', label: '자재투입현황', need: ['투입번호', '투입중량'], optional: true },
  { kw: 'FCO', label: 'FCO 수불', need: ['품목코드', '현재고', '현재고금액'], monthly: true, optional: true },
  { kw: '로트목록', label: '평택 로트목록', need: ['완료수량'], alias: ['당월생산'], lots: true },
];

function bookFromWorkbooks(XLSX, files) {
  const clean = s => String(s ?? '').replace(/\s/g, '');
  const monthIn = s => { const m = String(s || '').match(/(\d{1,2})\s*월/); return m ? Number(m[1]) : null; };
  const byName = s => SHEET_KINDS.find(k => [k.kw, ...(k.alias || [])].some(w => clean(s).includes(w)));
  const byHeader = rows => [...SHEET_KINDS].sort((a, b) => !!b.lots - !!a.lots).find(k => {
    for (let i = 0; i < Math.min(rows.length, 15); i++) {
      const c = (rows[i] || []).map(x => clean(x));
      if (k.lots ? c.includes('완료수량') && c.some(x => x.startsWith('로트번호')) : k.need.every(n => c.includes(n))) return true;
    }
    return false;
  });

  const sheets = [], recognized = [], have = new Set();
  const all = files.flatMap(f => sheetsOf(XLSX, f.wb).map(s => ({ file: f.name, sheet: s })));
  // 1차: 시트 이름으로 알 수 있는 것
  for (const x of all) {
    x.kind = byName(x.sheet.name);
    if (x.kind) { x.month = monthIn(x.sheet.name) ?? monthIn(x.file); x.how = '시트 이름'; }
  }
  // 2차: 머리글, 3차: 시트가 하나뿐인 파일의 파일 이름
  for (const x of all) {
    if (x.kind) continue;
    x.kind = byHeader(x.sheet.rows);
    if (x.kind) x.how = '머리글';
    else if (all.filter(y => y.file === x.file).length === 1 && byName(x.file)) { x.kind = byName(x.file); x.how = '파일 이름'; }
    if (x.kind) x.month = monthIn(x.sheet.name) ?? monthIn(x.file);
  }
  for (const x of all) {
    const k = x.kind;
    if (!k) { recognized.push({ file: x.file, sheet: x.sheet.name, kind: null, note: '사용 안 함' }); sheets.push(x.sheet); continue; }
    const id = k.kw + (k.many ? '|' + x.month : '');
    const rec = { file: x.file, sheet: x.sheet.name, kind: k.label, month: k.monthly ? x.month : null, how: x.how };
    if (have.has(id) && !k.lots) {
      rec.note = '같은 종류가 이미 있어 사용 안 함'; rec.dup = true;
      recognized.push(rec);
      continue;
    }
    have.add(id);
    if (k.many && x.month == null) rec.note = "월을 알 수 없음 — 파일·시트 이름에 '9월'처럼 넣어 주세요";
    const named = byName(x.sheet.name) === k;
    const name = named ? x.sheet.name : (k.monthly && x.month ? `${x.month}월 ` : '') + k.label.replace(/\s/g, ' ');
    sheets.push({ ...x.sheet, name, file: x.file });
    recognized.push(rec);
  }
  const missing = SHEET_KINDS.filter(k => !k.optional && !recognized.some(r => r.kind === k.label && !r.dup)).map(k => k.label);
  return { sheets, recognized, missing };
}


// ---------- 자료별 칸(슬롯) ----------
// 칸에 넣은 파일은 이름과 상관없이 그 칸의 용도로 씁니다.
// month: 'cur' = 당월, 'prev' = 전월
const SLOTS = [
  { id: 'lots', kind: '로트목록', label: '평택 로트목록', hint: '제품코드 · 로트번호 · 완료수량 · 생산공장 · 포장공장' },
  { id: 'ship', kind: '출하대장', label: '출하대장', month: 'cur', hint: '품목코드 · 로트번호 · 출하량 · 중량 · 공급가액' },
  { id: 'order', kind: '발주현황', label: '발주 현황', hint: '매출 월 · LOT · 제품코드 · 수주수량 · 수주단가' },
  { id: 'wo', kind: '작업지시', label: '작업지시현황', hint: '지시번호 · 제품로트번호 · 지시제품 (10-30 = 평택)' },
  { id: 'ledCur', kind: '단가대장', label: '제품 단가 대장', month: 'cur', hint: '품목코드 · 생산중량 · 자재코드 · 상세금액' },
  { id: 'ledPrev', kind: '단가대장', label: '제품 단가 대장', month: 'prev', optional: true, hint: '전월 생산분 출하 원가에 씁니다' },
  { id: 'ledPrev2', kind: '단가대장', label: '제품 단가 대장', month: 'prev2', optional: true, hint: '더 오래된 생산분이 출하될 때' },
  { id: 'fco', kind: 'FCO', label: 'FCO 수불', month: 'cur', optional: true, hint: '넣으면 미출고 기말 제품재고액을 FCO로 계산' },
];
const prevMonth = m => (m ? (m === 1 ? 12 : m - 1) : null);
const slotMonth = (slot, month) => (slot.month === 'cur' ? month : slot.month === 'prev' ? prevMonth(month) : slot.month === 'prev2' ? prevMonth(prevMonth(month)) : null);
const _clean = s => String(s ?? '').replace(/\s/g, '');
const _monthIn = s => { const m = String(s || '').match(/(\d{1,2})\s*월/); return m ? Number(m[1]) : null; };
const _kindOf = kw => SHEET_KINDS.find(k => k.kw === kw);
const _nameIs = (name, k) => [k.kw, ...(k.alias || [])].some(w => _clean(name).includes(w));

// 머리글에서 빠진 열 (없으면 빈 배열)
function headerMissing(rows, kw) {
  const k = _kindOf(kw);
  const need = k.lots ? ['로트번호', '완료수량'] : k.need;
  let best = need;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const c = (rows[i] || []).map(_clean);
    const miss = need.filter(n => !c.some(x => (k.lots ? x.startsWith(n) : x === n)));
    if (miss.length < best.length) best = miss;
  }
  return best;
}

// 통합문서에서 칸에 맞는 시트 고르기
// strict = true(통합 파일 자동 배치): 맞는 시트가 없으면 null
// strict = false(칸에 직접 넣음): 못 찾으면 첫 시트
function pickSheet(XLSX, wb, slot, month, strict) {
  const k = _kindOf(slot.kind), target = slotMonth(slot, month);
  const sheets = sheetsOf(XLSX, wb).filter(s => s.rows.some(r => r.some(v => v != null && v !== '')));
  const named = sheets.filter(s => _nameIs(s.name, k));
  let pick = null;
  if (slot.month && target) pick = named.find(s => _monthIn(s.name) === target);
  if (!pick && named.length) {
    if ((slot.month === 'prev' || slot.month === 'prev2') && strict) pick = null;            // 전월 칸은 월이 맞을 때만 자동 배치
    else pick = slot.month ? named.find(s => _monthIn(s.name) == null) || (strict ? null : named[0]) : named[0];
  }
  if (!pick && !(strict && named.length)) pick = sheets.find(s => !headerMissing(s.rows, slot.kind).length && !SHEET_KINDS.some(o => o !== k && _nameIs(s.name, o)));
  if (!pick && !strict) pick = sheets[0];
  return pick ? pick.name : null;
}

// filled = { 슬롯id: { wb, sheet, file } }
function bookFromSlots(XLSX, filled, month) {
  const pm = prevMonth(month);
  const nameFor = { order: '발주 현황', wo: '작업지시현황', ship: `${month}월 출하대장`, ledCur: `${month}월 제품 단가 대장`,
    ledPrev: `${pm}월 제품 단가 대장`, ledPrev2: `${prevMonth(pm)}월 제품 단가 대장`, fco: `${month}월 FCO`, lots: '평택 로트목록' };
  const sheets = [], warn = [];
  for (const slot of SLOTS) {
    const f = filled[slot.id];
    if (!f) continue;
    const sh = sheetsOf(XLSX, f.wb).find(s => s.name === f.sheet) || sheetsOf(XLSX, f.wb).find(s => s.name.trim() === String(f.sheet).trim());
    if (!sh) { warn.push({ slot: slot.id, missing: [`시트 「${f.sheet}」를 찾지 못함`] }); continue; }
    sheets.push({ ...sh, name: nameFor[slot.id], file: f.file, slot: slot.id });
    const miss = headerMissing(sh.rows, slot.kind);
    if (miss.length) warn.push({ slot: slot.id, missing: miss });
  }
  const missing = SLOTS.filter(s => !s.optional && !filled[s.id]).map(s => s.label + (s.month === 'cur' && s.kind === '단가대장' ? ' (당월)' : ''));
  return { sheets, missing, warn };
}

if (typeof module !== 'undefined') {
  module.exports = bookFromWorkbook;
  module.exports.bookFromWorkbooks = bookFromWorkbooks;
  module.exports.SHEET_KINDS = SHEET_KINDS;
  Object.assign(module.exports, { SLOTS, pickSheet, bookFromSlots, headerMissing, prevMonth, sheetsOf });
}

// SheetJS 통합문서 → 엔진 입력 (A1 기준 2차원 배열 + 수식 조회)
function sheetsOf(XLSX, wb) {
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
  { kw: '자재투입', label: '자재투입현황', need: ['투입번호', '투입중량'] },
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
  const missing = SHEET_KINDS.filter(k => !recognized.some(r => r.kind === k.label && !r.dup)).map(k => k.label);
  return { sheets, recognized, missing };
}

if (typeof module !== 'undefined') {
  module.exports = bookFromWorkbook;
  module.exports.bookFromWorkbooks = bookFromWorkbooks;
  module.exports.SHEET_KINDS = SHEET_KINDS;
}

// ===== 월별 자료 보관함: 미가공 파일 → 월별 저장 단위 (브라우저·Node 공용) =====
// 저장은 index.html이 브라우저(IndexedDB)에 하고, 여기서는 나누고 합치는 규칙만 둡니다.
// 원칙: 월은 데이터 안에서 정합니다. 추정하지 않습니다.
//  · 제품 단가 대장: 그 달 마감의 생산 투입액. 시트·파일 이름의 월(마감월)로 저장하고, 작업일자가 다른 행은 메모로 알림
//    (이름에 월이 없으면 작업일자 월로 나눔). 로트목록의 생산귀속도 마감월(예: 7월마감)이라 서로 맞춰짐
//  · 출하대장: 행마다 '출고일자'
//  · 평택 로트목록: 로트마다 '생산귀속 월'(예: 9월마감). 이월·취소·보류·예정은 저장하지 않음
//  · FCO: 시트·파일 이름의 월 (없으면 사용자가 월을 지정)
//  · 발주 현황·작업지시현황: 매달 갱신되는 누적 파일 → 최신본으로 덮어씀
const Archive = (() => {
  const isNode = typeof module !== 'undefined' && typeof window === 'undefined';
  const P = isNode ? require('./parse.js') : { sheetsOf, SHEET_KINDS };
  const E = isNode ? require('./engine.js') : Engine;
  const T = v => (v === null || v === undefined) ? '' : String(v).trim();
  const clean = s => T(s).replace(/\s/g, '');
  const monthIn = s => { const m = T(s).match(/(\d{1,2})\s*월/); return m ? Number(m[1]) : null; };
  const pad = m => String(m).padStart(2, '0');

  // 날짜 칸 → { y, m }  (Date, 2026.09 숫자, 2026.1(=10월) 숫자, '2026.09.23', '2026-09')
  function ymOf(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1 };
    if (typeof v === 'number') {
      if (v > 1900 && v < 2200) { const y = Math.floor(v); const m = Math.round((v - y) * 100); return m >= 1 && m <= 12 ? { y, m } : null; }
      return null;
    }
    const s = T(v).match(/^(\d{4})\s*[.\-/년]\s*(\d{1,2})/);
    if (s) { const m = Number(s[2]); return m >= 1 && m <= 12 ? { y: Number(s[1]), m } : null; }
    return null;
  }
  const key = (y, m) => `${y}-${pad(m)}`;
  // 연도가 없는 '9월' → 오늘 기준 가장 최근의 9월
  const ymFromMonth = (m, today) => { const y = today.getFullYear(); return key(m > today.getMonth() + 1 ? y - 1 : y, m); };
  const label = ym => `${Number(ym.slice(0, 4))}년 ${Number(ym.slice(5))}월`;

  function findHeader(rows, keyName) {
    for (let i = 0; i < Math.min(rows.length, 15); i++) if ((rows[i] || []).some(c => clean(c) === keyName)) return i;
    return -1;
  }
  // 필요한 열만 골라 [머리글, ...행]으로 (같은 이름이 여러 번이면 첫 번째)
  function pickCols(rows, hi, names) {
    const h = (rows[hi] || []).map(clean);
    const idx = names.map(n => h.indexOf(n));
    const out = [names];
    for (let i = hi + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const row = idx.map(j => (j >= 0 ? (r[j] instanceof Date ? r[j].toISOString().slice(0, 10) : r[j] ?? null) : null));
      if (row.every(v => v == null || v === '')) continue;
      out.push(row);
    }
    return out;
  }

  function kindOf(sheet, fileName, single) {
    const byName = s => P.SHEET_KINDS.find(k => [k.kw, ...(k.alias || [])].some(w => clean(s).includes(w)));
    let k = byName(sheet.name);
    if (!k) k = [...P.SHEET_KINDS].sort((a, b) => !!b.lots - !!a.lots).find(kk => {
      for (let i = 0; i < Math.min(sheet.rows.length, 15); i++) {
        const c = (sheet.rows[i] || []).map(clean);
        if (kk.lots ? c.includes('완료수량') && c.some(x => x.startsWith('로트번호')) : kk.need.every(n => c.includes(n))) return true;
      }
      return false;
    });
    if (!k && single) k = byName(fileName);
    return k || null;
  }

  const COLS = {
    led: ['작업일자', '품목코드', '품목명', '규격', '생산중량', '투입중량(원료)', '제조단가', '제조금액', '자재코드', '원료(자재)명', '사용중량', '자재단가', '상세금액'],
    ship: ['출고일자', '출하번호', '매출업체코드', '매출업체', '품목코드', '제품명', '규격', '로트번호', '소비기한', '출하량', '단위', '중량', '단가', '공급가액', '현장명칭'],
    fco: ['품목코드', '품목명', '단위', '전재고', '전재고액', '입고', '입고금액', '출고', '출고금액', '현재고', '현재고금액', '품목명칭', '규격', '매출수량', '매출금액', '대분류'],
    order: ['매출월', 'LOT', '제품코드', '품목', '구분', '규격', '수주수량', '수주단가', '출고일', '업체명'],
    wo: ['지시번호', '지시일자', '생산라인', '제품로트번호', '지시제품', '제품명', '규격', '지시량', '단위중량'],
  };
  const LOT_HDR = ['제품코드', '로트번호', '제품명', '고객사', '생산구분', '완료수량', '완료단가', '생산귀속월', '비고'];

  // 파일 하나 → { records, skipped, needMonth, notes }
  // record = { key: 'led:2026-09' | 'ship:…' | 'fco:…' | 'lots:…' | 'order' | 'wo', kind, ym, rows, n, src }
  // rules: { ledMonth: 'close'(시트 이름의 마감월) | 'date'(작업일자), exclude: ['이월', …] }
  function ingest(XLSX, wb, fileName, today = new Date(), rules = {}) { return ingestSheets(P.sheetsOf(XLSX, wb), fileName, today, rules); }
  function ingestSheets(allSheets, fileName, today = new Date(), rules = {}) {
    const ledMonth = rules.ledMonth || 'close';
    const exWords = rules.exclude && rules.exclude.length ? rules.exclude : ['이월', '취소', '보류', '예정'];
    const exRe = new RegExp(exWords.map(w => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'));
    const out = { records: [], skipped: [], needMonth: [], notes: [] };
    const sheets = allSheets.filter(s => s.rows.some(r => (r || []).some(v => v != null && v !== '')));
    const fileM = monthIn(fileName);
    for (const s of sheets) {
      const k = kindOf(s, fileName, sheets.length === 1);
      const src = `${fileName} · ${s.name.trim() || '(이름 없음)'}`;
      const sheetM = monthIn(s.name) ?? fileM;
      if (!k) { out.skipped.push({ src, reason: '종류를 알 수 없음' }); continue; }
      if (k.kw === '자재투입') { out.skipped.push({ src, reason: '자재투입현황은 계산에 쓰지 않음' }); continue; }
      if (k.kw === '단가대장' || k.kw === '출하대장') {
        const kind = k.kw === '단가대장' ? 'led' : 'ship';
        const hi = findHeader(s.rows, k.kw === '단가대장' ? '자재코드' : '로트번호');
        if (hi < 0) { out.skipped.push({ src, reason: '머리글을 찾지 못함' }); continue; }
        const tbl = pickCols(s.rows, hi, COLS[kind]);
        const dateCol = 0, groups = new Map(), noDate = [];
        for (const r of tbl.slice(1)) {
          if (kind === 'ship' && !T(r[1]) && !T(r[0])) continue;              // 합계 줄
          if (kind === 'ship' && clean(r[7]) === '로트번호') continue;          // 중간 머리글
          if (kind === 'led' && clean(r[1]) === '품목코드') continue;
          if (!T(r[1]) && kind === 'led') continue;
          const d = ymOf(r[dateCol]);
          if (!d) { noDate.push(r); continue; }
          const ym = key(d.y, d.m);
          if (!groups.has(ym)) groups.set(ym, []);
          groups.get(ym).push(r);
        }
        if (kind === 'led' && ledMonth === 'close' && sheetM && groups.size) {
          // 마감월 기준: 이름의 월로 한 번에 저장. 연도는 같은 달 작업일자(없으면 가장 많은 연도)에서
          const same = [...groups.keys()].find(ym => Number(ym.slice(5)) === sheetM);
          const yr = same ? same.slice(0, 4) : [...groups.entries()].sort((a, b) => b[1].length - a[1].length)[0][0].slice(0, 4);
          const ym = `${yr}-${pad(sheetM)}`;
          const rows = [...groups.values()].flat();
          out.records.push({ key: `led:${ym}`, kind, ym, rows: [COLS.led, ...rows], n: rows.length, src });
          const other = [...groups.entries()].filter(([g]) => g !== ym);
          if (other.length) {
            const prods = new Set(other.flatMap(([, rs]) => rs.map(r => T(r[1])).filter(p => p.startsWith('1'))));
            out.notes.push(`${src}: 작업일자가 ${other.map(([g, rs]) => `${label(g)} ${rs.length}행`).join(', ')}인 행이 있음(제품 ${[...prods].join(', ') || '없음'}) → ${label(ym)} 마감 단가대장으로 함께 저장`);
          }
        } else {
          for (const [ym, rows] of groups) out.records.push({ key: `${kind}:${ym}`, kind, ym, rows: [COLS[kind], ...rows], n: rows.length, src });
          if (groups.size > 1) out.notes.push(`${src}: ${[...groups.keys()].map(label).join(', ')}이 섞여 있어 ${kind === 'led' ? '작업일자' : '출고일자'} 기준으로 달마다 나눠 저장`);
          if (sheetM && groups.size === 1) { const only = [...groups.keys()][0]; if (Number(only.slice(5)) !== sheetM) out.notes.push(`${src}: 시트·파일 이름은 ${sheetM}월인데 출고일자는 ${label(only)} → 출고일자의 월로 저장`); }
        }
        if (noDate.length) out.notes.push(`${src}: ${kind === 'led' ? '작업일자' : '출고일자'}가 없는 ${noDate.length}행은 저장하지 않음`);
        continue;
      }
      if (k.kw === 'FCO') {
        const hi = findHeader(s.rows, '현재고금액');
        if (hi < 0) { out.skipped.push({ src, reason: '머리글을 찾지 못함' }); continue; }
        const tbl = pickCols(s.rows, hi, COLS.fco);
        const rows = tbl.slice(1).filter(r => T(r[0]) && clean(r[0]) !== '품목코드');
        const rec = { key: null, kind: 'fco', ym: null, rows: [COLS.fco, ...rows], n: rows.length, src };
        if (sheetM) { rec.ym = ymFromMonth(sheetM, today); rec.key = `fco:${rec.ym}`; out.records.push(rec); }
        else out.needMonth.push(rec);
        continue;
      }
      if (k.kw === '발주현황' || k.kw === '작업지시') {
        const kind = k.kw === '발주현황' ? 'order' : 'wo';
        const hi = findHeader(s.rows, kind === 'order' ? 'LOT' : '제품로트번호');
        if (hi < 0) { out.skipped.push({ src, reason: '머리글을 찾지 못함' }); continue; }
        const tbl = pickCols(s.rows, hi, COLS[kind]);
        const rows = tbl.slice(1).filter(r => kind === 'order' ? T(r[1]) && clean(r[1]) !== 'LOT' : T(r[4]) && clean(r[4]) !== '지시제품');
        out.records.push({ key: kind, kind, ym: null, rows: [COLS[kind], ...rows], n: rows.length, src });
        continue;
      }
      if (k.lots) {
        let lots;
        try { lots = E.readLots({ sheets: [s] }).lots; } catch (e) { out.skipped.push({ src, reason: '로트목록 열을 찾지 못함' }); continue; }
        const groups = new Map(), noMonth = [], skip = [];
        for (const l of lots) {
          const g = T(l.생산귀속);
          if (exRe.test(g)) { skip.push(`${l.로트번호}(${g})`); continue; }
          let m = monthIn(g);
          if (!m && l.구분 === '전월' && sheetM) m = sheetM === 1 ? 12 : sheetM - 1;
          const row = [l.제품코드, l.로트번호, l.제품명, l.고객사, l.생산구분, l.완료수량, l.완료단가, g, ''];
          if (/^10[A-Z]/.test(l.제품코드)) row[8] = '제품코드 확인(O를 0으로 입력한 것으로 보임)';
          if (!m) { noMonth.push(row); continue; }
          const ym = ymFromMonth(m, today);
          if (!groups.has(ym)) groups.set(ym, []);
          groups.get(ym).push(row);
        }
        for (const [ym, rows] of groups) out.records.push({ key: `lots:${ym}`, kind: 'lots', ym, rows: [LOT_HDR, ...rows], n: rows.length, src });
        if (skip.length) out.notes.push(`${src}: 생산귀속이 ${exWords.join('·')}인 ${skip.length}개는 저장하지 않음 (${skip.slice(0, 5).join(', ')}${skip.length > 5 ? ' 외' : ''})`);
        const typo = [...groups.values()].flat().filter(r => r[8]).map(r => r[0]);
        if (typo.length) out.notes.push(`${src}: 제품코드 형식 확인 필요 ${[...new Set(typo)].join(', ')} (O 대신 0?) — 고치지 않고 그대로 저장`);
        if (noMonth.length) out.needMonth.push({ key: null, kind: 'lots', ym: null, rows: [LOT_HDR, ...noMonth], n: noMonth.length, src });
        continue;
      }
      out.skipped.push({ src, reason: '종류를 알 수 없음' });
    }
    return out;
  }

  // 같은 자리에 저장할 때: 로트목록은 제품코드+로트로 합치고(새 것 우선), 나머지는 새 것으로 바꿈
  function merge(oldRec, newRec) {
    if (!oldRec || newRec.kind !== 'lots') return { ...newRec, savedAt: new Date().toISOString() };
    const ok = r => ['평택', '화성포장', '멀티연질', '정제화성', '화성'].includes(T(r[4])) && r[5] != null;
    const map = new Map(oldRec.rows.slice(1).map(r => [r[0] + '|' + r[1], r]));
    const kept = [];
    newRec.rows.slice(1).forEach(r => { const k = r[0] + '|' + r[1], o = map.get(k); if (o && ok(o) && !ok(r)) kept.push(r[1]); else map.set(k, r); });
    if (kept.length) newRec = { ...newRec, kept };
    const rows = [...map.values()];
    return { ...newRec, rows: [LOT_HDR, ...rows], n: rows.length, src: oldRec.src === newRec.src ? newRec.src : `${oldRec.src} + ${newRec.src}`, savedAt: new Date().toISOString() };
  }

  // 보관함 자료 고치기: 다른 월로 옮기기, 로트목록 행 수정. 수정 이력(edits)을 남김
  function moveRecord(rec, ym, why) {
    const edits = [...(rec.edits || []), { at: new Date().toISOString(), what: `${label(rec.ym)} → ${label(ym)} 이동${why ? ' (' + why + ')' : ''}` }];
    return { ...rec, ym, key: `${rec.kind}:${ym}`, edits, savedAt: new Date().toISOString() };
  }
  function editLots(rec, rows, why) {
    const before = new Map(rec.rows.slice(1).map(r => [r[1] + '#' + rec.rows.indexOf(r), r]));
    const changes = [];
    rows.forEach((r, i) => { const o = rec.rows[i + 1]; if (!o) return; LOT_HDR.forEach((h, j) => { if (String(o[j] ?? '') !== String(r[j] ?? '')) changes.push(`${o[1]} ${h}: ${o[j] ?? ''} → ${r[j] ?? ''}`); }); });
    const edits = changes.length ? [...(rec.edits || []), { at: new Date().toISOString(), what: changes.join('; ') + (why ? ' (' + why + ')' : '') }] : (rec.edits || []);
    return { ...rec, rows: [LOT_HDR, ...rows], n: rows.length, edits, savedAt: new Date().toISOString() };
  }

  // 보관함 → 계산용 book (당월 ym) + 생산월 기록
  // all = { [key]: record }
  function buildBook(all, ym) {
    const m = Number(ym.slice(5)), y = Number(ym.slice(0, 4));
    const within = k => { const a = Number(k.slice(0, 4)) * 12 + Number(k.slice(5)), b = y * 12 + m; return a <= b && a > b - 12; };
    const sheet = (name, rows) => ({ name, rows, formula: () => null });
    const sheets = [], missing = [];
    const lots = all[`lots:${ym}`]; lots ? sheets.push(sheet('평택 로트목록', lots.rows)) : missing.push('평택 로트목록');
    const ship = all[`ship:${ym}`]; ship ? sheets.push(sheet(`${m}월 출하대장`, ship.rows)) : missing.push('출하대장');
    all.order ? sheets.push(sheet('발주 현황', all.order.rows)) : missing.push('발주 현황');
    all.wo ? sheets.push(sheet('작업지시현황', all.wo.rows)) : missing.push('작업지시현황');
    const leds = Object.values(all).filter(r => r.kind === 'led' && within(r.ym));
    if (!all[`led:${ym}`]) missing.push('제품 단가 대장(당월)');
    leds.forEach(r => sheets.push(sheet(`${Number(r.ym.slice(5))}월 제품 단가 대장`, r.rows)));
    const fco = all[`fco:${ym}`]; if (fco) sheets.push(sheet(`${m}월 FCO`, fco.rows));
    // 생산월 기록: 이전 달 로트목록의 제품코드+로트 → 생산귀속 월
    const lotMonth = new Map(), conflict = [];
    Object.values(all).filter(r => r.kind === 'lots' && within(r.ym) && r.ym !== ym).sort((a, b) => a.ym < b.ym ? -1 : 1).forEach(r => {
      const mm = Number(r.ym.slice(5));
      r.rows.slice(1).forEach(x => { const k = x[0] + '|' + x[1]; if (lotMonth.has(k) && lotMonth.get(k) !== mm) conflict.push(k); else lotMonth.set(k, mm); });
    });
    return { book: { sheets }, history: { lotMonth, conflict }, missing, ledMonths: leds.map(r => r.ym).sort() };
  }

  return { ingest, ingestSheets, merge, buildBook, moveRecord, editLots, ymOf, label, ymFromMonth, LOT_HDR };
})();
if (typeof module !== 'undefined' && typeof window === 'undefined') module.exports = Archive;

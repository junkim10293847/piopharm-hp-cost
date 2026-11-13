// ===== 화성·평택 매출/원가 계산 엔진 v2 =====
// 2026-11 개편: 2026년 9월 가공 파일과 원 단위 대조로 검증한 방식
//  · 로트는 항상 제품코드 + 로트번호로 찾음 (로트 이름을 고치지 않아도 됨)
//  · 원가는 출하된 수량만: 출하량 × (당월 단가대장 제조금액 ÷ 그 제품 로트들의 생산수량 합)
//  · 조정비율·규격 해석·자재투입 로트 조회는 쓰지 않음
//  · 이전 생산분은 생산월 단가대장의 kg당 원가 × 출하중량. 생산월은 로트목록 생산귀속(당월 목록·보관함의 이전 달 목록)으로만 정함
//    생산월이나 그 달 대장이 없으면 원가 빈칸 (작업지시 날짜 등으로 추정하지 않음)
//  · 로트목록에 없는 평택 출하(작업지시 10-30 또는 비오팜2공장)는 '전월 생산 출하'로 따로 계산
// book = { sheets: [ { name, rows: any[][] (A1 기준), formula(r,c) } ] }
// opts = { month, ratio: {생산구분: 비율} }
const Engine = (() => {
  const T = v => (v === null || v === undefined) ? '' : (v instanceof Date ? v.toISOString() : String(v).trim());
  const N = v => {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'boolean' || v instanceof Date) return null;
    const x = Number(String(v).replace(/[,\s]/g, ''));
    return isNaN(x) ? null : x;
  };
  const Z = v => v == null ? 0 : v;
  const Clean = s => T(s).replace(/[\s]/g, '');
  const noSpace = s => s.replace(/\s/g, '');
  const RATIO = { '평택': 1, '화성포장': 0.8, '멀티연질': 0.6, '정제화성': 0.2, '화성': 0 };
  // 바뀔 수 있는 규칙 (화면 「계산 규칙」에서 바꿈). run()이 시작할 때 정해짐
  const DEFAULT_RULES = { ptWo: '10-30', ptSite: '비오팜2공장', semiPrefix: '3', alcohol: ['4C00000149'], exclude: ['이월', '취소', '보류', '예정'] };
  let RU = { ...DEFAULT_RULES };
  const exRe = () => new RegExp((RU.exclude && RU.exclude.length ? RU.exclude : ['\\uFFFF']).map(w => w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|'));

  function sheetBy(book, kw, optional) {
    const s = book.sheets.find(x => noSpace(x.name).includes(kw));
    if (!s && !optional) throw new Error(`원천 파일에 '${kw}' 시트가 없습니다`);
    return s || null;
  }
  function promote(rows, key) {
    const lim = Math.min(rows.length, 15);
    for (let i = 0; i < lim; i++) {
      if ((rows[i] || []).some(c => Clean(c) === key)) return { hdr: rows[i].map(Clean), data: rows.slice(i + 1), at: i };
    }
    throw new Error(`머리글을 찾지 못했습니다: ${key}`);
  }
  // keys: 꼭 필요한 열, opt: 없어도 되는 열
  function pick(pr, keys, opt = []) {
    const pos = keys.map(k => {
      const p = pr.hdr.indexOf(k);
      if (p < 0) throw new Error(`열을 찾지 못했습니다: ${k}`);
      return p;
    });
    const opos = opt.map(k => pr.hdr.indexOf(k));
    return pr.data.map((r, i) => {
      const o = { _row: pr.at + i + 2 };
      keys.forEach((k, j) => o[k] = (r || [])[pos[j]] ?? null);
      opt.forEach((k, j) => o[k] = opos[j] >= 0 ? (r || [])[opos[j]] ?? null : null);
      return o;
    });
  }
  const sheetMonth = name => { const m = name.match(/(\d{1,2})\s*월/); return m ? Number(m[1]) : null; };
  const prevM = m => (m === 1 ? 12 : m - 1);

  // ---------- 로트목록 ----------
  // 지원 형식: ① 미가공 「9월」 시트(생산공장·포장공장) ② 기존 「당월 생산」(평택비율 텍스트, 전월 구간) ③ 「로트목록」 표
  function readLots(book) {
    for (const s of book.sheets) {
      const rows = s.rows;
      const hdrIdx = [];
      rows.forEach((r, i) => {
        const c = (r || []).map(Clean);
        if (c.includes('완료수량') && c.some(x => x.startsWith('로트번호')) && c.some(x => x.startsWith('제품코드'))) hdrIdx.push(i);
      });
      if (!hdrIdx.length) continue;
      const lots = [];
      hdrIdx.forEach((hi, k) => {
        const h = rows[hi].map(Clean);
        const find = (...names) => { for (const n of names) { const p = h.indexOf(n); if (p >= 0) return p; } return -1; };
        const findStart = pre => h.findIndex(x => x.startsWith(pre));
        const cProd = find('제품코드(제공자료)') >= 0 ? find('제품코드(제공자료)') : findStart('제품코드');
        const cLot = find('로트번호(제공자료)', '로트번호', '로트번호(확인)');
        const cName = find('제품명'), cCust = find('고객사');
        const cLine = find('생산구분', '평택비율');
        const cMake = find('생산공장'), cPack = find('포장공장');
        const cQty = find('완료수량'), cPrice = find('완료단가') >= 0 ? find('완료단가') : find('단가');
        const cMonth = find('(생산귀속월)', '생산귀속월');
        let section = '당월';
        for (let j = Math.max(0, hi - 3); j < hi; j++) if ((rows[j] || []).some(c => T(c).includes('전월'))) section = '전월';
        const end = k + 1 < hdrIdx.length ? hdrIdx[k + 1] : rows.length;
        for (let i = hi + 1; i < end; i++) {
          const r = rows[i] || [];
          const lot = T(r[cLot]); const qty = N(r[cQty]);
          if (!lot || qty == null || cProd < 0) continue;
          let line = cLine >= 0 ? T(r[cLine]) : '';
          if (!RATIO.hasOwnProperty(line)) {
            const mk = cMake >= 0 ? T(r[cMake]) : '', pk = cPack >= 0 ? T(r[cPack]) : '';
            if (mk === '평택' && pk === '평택') line = '평택';
            else if (mk === '평택' && pk === '화성') line = '화성포장';
            else if (mk === '화성' && pk === '화성') line = '화성';
            else if (!line) line = mk || pk ? `${mk}/${pk}` : '';
          }
          lots.push({
            구분: section, 행: i + 1, 제품코드: T(r[cProd]), 로트번호: lot, 제품명: T(r[cName]), 고객사: T(r[cCust]),
            생산구분: line, 완료수량: qty, 완료단가: N(r[cPrice]), 생산귀속: cMonth >= 0 ? T(r[cMonth]) : '',
          });
        }
      });
      if (lots.length) return { sheet: s.name, lots };
    }
    throw new Error("로트목록을 찾지 못했습니다. 평택 로트목록(제품코드·로트번호·완료수량 열)이 필요합니다.");
  }

  function detectMonth(book) {
    const s = book.sheets.find(x => noSpace(x.name).includes('출하대장'));
    const m = s && sheetMonth(s.name);
    if (m) return m;
    const ms = book.sheets.filter(x => noSpace(x.name).includes('단가대장')).map(x => sheetMonth(x.name)).filter(Boolean);
    return ms.length ? Math.max(...ms) : null;
  }

  // ---------- 원천 시트 준비 ----------
  function prepare(book, month) {
    const OdAll = pick(promote(sheetBy(book, '발주현황').rows, 'LOT'), ['매출월', 'LOT', '제품코드', '구분', '수주수량', '수주단가'], ['품목', '출고일'])
      .filter(r => T(r.LOT) && T(r.LOT) !== 'LOT')
      .map(r => ({ ...r, 월: N(r['매출월']), key: T(r['제품코드']) + '|' + T(r.LOT) }));
    const Od = OdAll.filter(r => r.월 === month);
    const Wo = pick(promote(sheetBy(book, '작업지시').rows, '제품로트번호'), ['제품로트번호', '지시제품'], ['지시번호', '지시일자', '지시량', '단위중량', '규격', '생산라인'])
      .filter(r => T(r['지시제품']))
      .map(r => ({ ...r, key: T(r['지시제품']) + '|' + T(r['제품로트번호']) }));
    const Sh = pick(promote(sheetBy(book, '출하대장').rows, '로트번호'), ['로트번호', '품목코드', '제품명', '매출업체', '출하량', '중량', '공급가액'], ['출고일자', '출하번호', '현장명칭'])
      .filter(r => T(r['품목코드']) && T(r['품목코드']) !== '품목코드' && (T(r['출하번호']) || r['출하번호'] == null))
      .filter(r => !(r['출하번호'] == null && r['출고일자'] == null && r['매출업체'] == null)) // 합계 줄 제외
      .map(r => ({ ...r, key: T(r['품목코드']) + '|' + T(r['로트번호']), qty: Z(N(r['출하량'])), w: Z(N(r['중량'])), amt: Z(N(r['공급가액'])) }));
    // 단가대장: 제품(1…) 행만, 제품·월별 합계
    const Lg = new Map();
    const ledSheets = book.sheets.filter(x => noSpace(x.name).includes('단가대장'));
    if (!ledSheets.length) throw new Error("원천 파일에 '제품 단가 대장' 시트가 없습니다");
    for (const s of ledSheets) {
      const m = sheetMonth(s.name);
      for (const r of pick(promote(s.rows, '자재코드'), ['품목코드', '자재코드', '상세금액', '생산중량'], ['품목명', '제조금액'])) {
        const p = T(r['품목코드']); if (!p.startsWith('1')) continue;
        const mat = T(r['자재코드']); const amt = Z(N(r['상세금액']));
        const k = p + '|' + m;
        if (!Lg.has(k)) Lg.set(k, { p, m, name: T(r['품목명']), 제조금액: 0, 원료: 0, pw: N(r['생산중량']), 행수: 0, 시트: s.name });
        const g = Lg.get(k);
        g.제조금액 += amt; g.행수++;
        if ((RU.semiPrefix && mat.startsWith(RU.semiPrefix)) || RU.alcohol.includes(mat)) g.원료 += amt;
      }
    }
    const led = (p, m) => Lg.get(p + '|' + m) || null;
    // 이전 생산분의 생산월: 로트목록의 생산귀속 월(당월 목록 또는 보관함의 이전 달 목록)로만 정함. 추정하지 않음
    const prodMonth = (lot, hist) => {
      if (lot) {
        const g = T(lot.생산귀속); const mm = g.match(/(\d{1,2})\s*월/);
        if (mm && !exRe().test(g)) return { m: Number(mm[1]), by: `로트목록 생산귀속 「${g}」` };
        if (lot.구분 === '전월') return { m: month === 1 ? 12 : month - 1, by: '로트목록 전월 구간' };
      }
      if (hist && hist.lotMonth && lot && hist.lotMonth.has(lot.key)) return { m: hist.lotMonth.get(lot.key), by: '보관함 이전 달 로트목록' };
      return null;
    };
    // FCO (선택)
    let Fco = null;
    const fs = sheetBy(book, 'FCO', true);
    if (fs) {
      Fco = new Map();
      for (const r of pick(promote(fs.rows, '현재고금액'), ['품목코드', '현재고', '현재고금액'], ['품목명칭', '품목명', '규격', '전재고', '전재고액', '입고', '입고금액', '출고', '출고금액', '대분류'])) {
        const p = T(r['품목코드']); if (!p) continue;
        const nums = ['전재고', '전재고액', '입고', '입고금액', '출고', '출고금액', '현재고', '현재고금액'];
        const o = { 품목코드: p, 품목명: T(r['품목명칭']) || T(r['품목명']), 규격: T(r['규격']), 대분류: T(r['대분류']), _row: r._row };
        nums.forEach(k => o[k] = Z(N(r[k])));
        Fco.set(p, o);
      }
    }
    return { OdAll, Od, Wo, Sh, Lg, led, prodMonth, Fco };
  }

  const isPtWo = w => !!RU.ptWo && T(w['지시번호']).startsWith(RU.ptWo);
  // 출하대장 제품명이 숫자 코드(수출품 등)이면 FCO·단가대장 이름으로 바꿈
  const nameOf = (src, p, n) => {
    if (n && !/^\d+$/.test(n)) return n;
    const f = src.Fco && src.Fco.get(p); if (f && f.품목명) return f.품목명;
    const g = [...src.Lg.values()].find(x => x.p === p && x.name && !/^\d+$/.test(x.name));
    return g ? g.name : n;
  };

  // ---------- 평택 로트목록 로트 ----------
  function calcLots(lots, src, opts) {
    const { Od, OdAll, Wo, Sh, led, prodMonth } = src;
    const base = lots.map(l => {
      const key = l.제품코드 + '|' + l.로트번호;
      const K = RATIO.hasOwnProperty(l.생산구분) ? (opts.ratio[l.생산구분] ?? RATIO[l.생산구분]) : null;
      const sh = Sh.filter(x => x.key === key);
      const shipQty = sh.reduce((a, x) => a + x.qty, 0), shipW = sh.reduce((a, x) => a + x.w, 0), shipAmt = sh.reduce((a, x) => a + x.amt, 0);
      const cur = led(l.제품코드, opts.month);
      const pm = cur ? null : prodMonth({ ...l, key }, opts.history);
      const prev = pm && pm.m !== opts.month ? led(l.제품코드, pm.m) : null;
      return { ...l, key, K, sh, shipQty, shipW, shipAmt, cur, prev, 생산월: pm,
        wo: Wo.filter(x => x.key === key), od: Od.filter(x => x.key === key), odAll: OdAll.filter(x => x.key === key),
        생산수량: Math.max(Z(l.완료수량), shipQty) };
    });
    return base.map(r => {
      const o = { ...r, flags: [] };
      if (r.K == null) o.flags.push('생산구분 확인');
      if (!r.wo.length) o.flags.push('작업지시 없음');
      if (r.shipQty > Z(r.완료수량)) o.flags.push(`출하량(${r.shipQty.toLocaleString()}) > 완료수량`);
      if (r.cur) {
        // 당월 생산: 같은 제품의 당월 로트끼리 생산수량 비율
        const grp = base.filter(x => x.제품코드 === r.제품코드 && x.cur);
        const tot = grp.reduce((a, x) => a + x.생산수량, 0);
        o.원가월 = opts.month; o.원가방식 = '당월 생산'; o.제품생산합 = tot;
        o.개당원가 = tot ? r.cur.제조금액 / tot : 0;
        o.Full원가 = r.shipQty * o.개당원가;
        o.Full원료 = r.cur.제조금액 ? o.Full원가 * r.cur.원료 / r.cur.제조금액 : 0;
        // 단가대장 생산중량 vs 로트 수량 점검 (출하 kg/개 또는 작업지시 단위중량)
        const perUnit = x => (x.shipQty && x.shipW) ? x.shipW / x.shipQty : (x.wo.length ? N(x.wo[0]['단위중량']) : null);
        const lw = grp.reduce((a, x) => { const u = perUnit(x); return u == null ? NaN : a + u * x.생산수량; }, 0);
        o.단가대장중량 = r.cur.pw; o.로트중량합 = isNaN(lw) ? null : lw;
        if (r.cur.pw && o.로트중량합 && Math.abs(o.로트중량합 / r.cur.pw - 1) > 0.03) o.flags.push('단가대장 생산량 ≠ 로트 합계');
        if (!r.cur.제조금액) o.flags.push('단가대장 금액 0(임가공)');
        if (!r.shipQty) o.flags.push('출하 없음(당월 생산 → 미출고 재고)');
        else if (r.생산수량 > r.shipQty) o.flags.push(`부분 출하(잔량 ${(r.생산수량 - r.shipQty).toLocaleString()})`);
      } else if (r.shipQty) {
        o.원가방식 = '이전 생산(출하중량 × 생산월 kg당 원가)'; o.생산월근거 = r.생산월 ? r.생산월.by : null;
        if (r.prev && r.prev.pw) {
          o.원가월 = r.prev.m; o.개당원가 = r.shipQty ? (r.shipW * r.prev.제조금액 / r.prev.pw) / r.shipQty : 0;
          o.Full원가 = r.shipW * r.prev.제조금액 / r.prev.pw;
          o.Full원료 = r.prev.제조금액 ? o.Full원가 * r.prev.원료 / r.prev.제조금액 : 0;
        } else { o.원가월 = null; o.Full원가 = null; o.Full원료 = null; o.flags.push(r.생산월 ? `생산월 ${r.생산월.m}월 단가대장 없음 → 원가 빈칸` : '생산월 확인 불가(로트목록 생산귀속 없음) → 원가 빈칸'); }
      } else {
        o.원가방식 = '출하 없음'; o.원가월 = null; o.Full원가 = 0; o.Full원료 = 0;
        const om = [...new Set(r.odAll.map(x => x.월).filter(m => m != null))];
        o.flags.push(om.length && !om.includes(opts.month) ? `출하 없음 · 수주 매출월 ${om.join(',')}월(선매출·이월 의심)` : '출하 없음');
      }
      const K = o.K;
      o.매출 = K == null ? null : r.shipAmt * K;
      o.원가 = (K == null || o.Full원가 == null) ? null : o.Full원가 * K;
      o.원료 = (K == null || o.Full원료 == null) ? null : o.Full원료 * K;
      o.자재 = o.원가 == null ? null : o.원가 - o.원료;
      o.원가율 = (o.매출 && o.원가 != null) ? o.원가 / o.매출 : null;
      const h = K == null ? 0 : 1 - K;
      o.H매출 = r.shipAmt * h; o.H원료 = Z(o.Full원료) * h; o.H자재 = (Z(o.Full원가) - Z(o.Full원료)) * h;
      o.잔량 = Math.max(0, r.생산수량 - r.shipQty);
      o.잔량재료비 = r.cur ? o.잔량 * o.개당원가 : null;
      o.수주수량 = r.od.reduce((a, x) => a + Z(N(x['수주수량'])), 0);
      o.수주단가 = r.od.length ? N(r.od[0]['수주단가']) : null;
      if (r.shipQty && !r.od.length) o.flags.push('당월 수주 없음');
      return o;
    });
  }

  // ---------- 로트목록에 없는 평택 출하 (전월 이전 생산분) ----------
  function calcPrev(lotRows, src, opts) {
    const listed = new Set(lotRows.map(l => l.key));
    const groups = new Map();
    for (const r of src.Sh) {
      if (listed.has(r.key) || !T(r['품목코드']).startsWith('1')) continue;
      const wo = src.Wo.filter(x => x.key === r.key);
      const pt = wo.some(isPtWo) || T(r['현장명칭']) === RU.ptSite;
      if (!pt) continue;
      if (!groups.has(r.key)) groups.set(r.key, { key: r.key, 제품코드: T(r['품목코드']), 로트번호: T(r['로트번호']), 제품명: nameOf(src, T(r['품목코드']), T(r['제품명'])), 고객사: T(r['매출업체']), sh: [], wo });
      groups.get(r.key).sh.push(r);
    }
    return [...groups.values()].map(g => {
      const shipQty = g.sh.reduce((a, x) => a + x.qty, 0), shipW = g.sh.reduce((a, x) => a + x.w, 0), shipAmt = g.sh.reduce((a, x) => a + x.amt, 0);
      const pm = src.prodMonth({ key: g.key }, opts.history);
      const c = pm && pm.m !== opts.month ? src.led(g.제품코드, pm.m) : null;
      const o = { ...g, 구분: '전월 생산', K: 1, 생산구분: '평택', shipQty, shipW, shipAmt, flags: [], od: src.Od.filter(x => x.key === g.key), 생산월: pm, 생산월근거: pm ? pm.by : null };
      o.근거 = [g.wo.some(isPtWo) ? `작업지시 ${RU.ptWo}` : null, g.sh.some(x => T(x['현장명칭']) === RU.ptSite) ? RU.ptSite : null].filter(Boolean).join(' · ');
      o.지시일자 = g.wo.map(w => T(w['지시일자'])).filter(Boolean).join(', ');
      if (c && c.pw && shipW) {
        o.원가월 = c.m; o.원가 = shipW * c.제조금액 / c.pw; o.원료 = c.제조금액 ? o.원가 * c.원료 / c.제조금액 : 0; o.자재 = o.원가 - o.원료;
      } else { o.원가월 = null; o.원가 = null; o.원료 = null; o.자재 = null; o.flags.push(pm ? `생산월 ${pm.m}월 단가대장 없음 → 원가 빈칸` : '생산월 확인 불가(이전 달 로트목록에 없음) → 원가 빈칸');
        if (!pm && opts.history && opts.history.lotMonth) { // 제품코드 오타(O ↔ 0) 후보만 알려 주고 쓰지는 않음
          const alt = g.제품코드.replace(/^1O/, '10') + '|' + g.로트번호;
          if (opts.history.lotMonth.has(alt)) o.flags.push(`이전 로트목록에 ${alt.split('|')[0]}(제품코드 오타 의심)로 ${opts.history.lotMonth.get(alt)}월 기록 있음 → 원본 확인`);
        } }
      o.매출 = shipAmt; o.원가율 = (o.원가 != null && shipAmt) ? o.원가 / shipAmt : null;
      o.cur = c;
      return o;
    }).sort((a, b) => b.매출 - a.매출);
  }

  // ---------- 화성 품목별 (평택에 들어가지 않은 출하) ----------
  function calcHwaseong(taken, src, opts) {
    const groups = new Map();
    for (const r of src.Sh) {
      if (taken.has(r.key)) continue;
      const p = T(r['품목코드']);
      if (!groups.has(p)) groups.set(p, { 품목코드: p, 제품명: nameOf(src, p, T(r['제품명'])), 거래처: new Set(), 출하량: 0, 중량: 0, 매출: 0, 건수: 0, sh: [] });
      const g = groups.get(p);
      g.거래처.add(T(r['매출업체'])); g.출하량 += r.qty; g.중량 += r.w; g.매출 += r.amt; g.건수++; g.sh.push(r);
    }
    const out = [];
    for (const g of groups.values()) {
      const c = src.led(g.품목코드, opts.month);
      let unit = null, cost = null, raw = null, mat = null, flag = null;
      if (!c) flag = g.품목코드.startsWith('1') ? '당월 단가대장 없음 → 원가 빈칸' : '제품 외 품목(원가 대상 아님)';
      else if (!c.pw) flag = '단가대장 생산중량 없음';
      else if (!g.중량) flag = '출하중량 없음';
      else { unit = c.제조금액 / c.pw; cost = g.중량 * unit; raw = c.제조금액 ? cost * c.원료 / c.제조금액 : 0; mat = cost - raw; }
      out.push({ ...g, 거래처: [...g.거래처].filter(Boolean).join(', '), 원가월: c ? c.m : null, kg당원가: unit, cur: c,
        원료: raw, 자재: mat, 원가: cost, 원가율: (cost != null && g.매출) ? cost / g.매출 : null, flag });
    }
    return out.sort((a, b) => b.매출 - a.매출);
  }

  // ---------- 미출고 재고 (FCO가 있으면 FCO, 없으면 단가대장 기준 추정) ----------
  function calcStock(lots, src, opts) {
    const ptProd = new Set(lots.map(l => l.제품코드));
    const woByProd = new Map();
    for (const w of src.Wo) { const p = T(w['지시제품']); if (!woByProd.has(p)) woByProd.set(p, []); woByProd.get(p).push(w); }
    const plantOf = p => {
      if (ptProd.has(p)) return '평택';
      const ws = (woByProd.get(p) || []).filter(w => T(w['지시번호']));
      if (!ws.length) return src.Sh.some(x => T(x['품목코드']) === p && T(x['현장명칭']) === RU.ptSite) ? '평택' : '확인';
      const pt = ws.filter(isPtWo).length;
      return pt === ws.length ? '평택' : pt === 0 ? '화성' : '평택·화성';
    };
    const rows = [];
    // 평택 로트목록 중 당월 생산 잔량 (재료비 추정)
    const est = new Map();
    lots.filter(o => o.cur && o.잔량 > 0).forEach(o => {
      const e = est.get(o.제품코드) || { 수량: 0, 재료비: 0, 로트: [] };
      e.수량 += o.잔량; e.재료비 += Z(o.잔량재료비); e.로트.push(`${o.로트번호}(${o.잔량.toLocaleString()})`); est.set(o.제품코드, e);
    });
    if (src.Fco) {
      for (const f of src.Fco.values()) {
        if (!f.품목코드.startsWith('1') || !(f.현재고 > 0)) continue;
        const e = est.get(f.품목코드);
        const kind = f.입고 > 0 ? (f.출고 > 0 ? '당월 생산·일부 출고' : '당월 생산·미출고') : (f.출고 > 0 ? '이전 생산·일부 출고' : '이전 생산·움직임 없음');
        const unit = f.현재고 ? f.현재고금액 / f.현재고 : 0;
        const flags = [];
        if (f.현재고 < 100 && f.현재고금액 >= 1e6) flags.push('수량 대비 금액 큼(단위·단가 확인)');
        if (f.입고 > 0 && !src.led(f.품목코드, opts.month)) flags.push('당월 입고인데 단가대장 없음');
        rows.push({ 제품코드: f.품목코드, 제품명: f.품목명, 규격: f.규격, 공장: plantOf(f.품목코드), 구분: kind,
          전재고: f.전재고, 입고: f.입고, 출고: f.출고, 현재고: f.현재고, 현재고금액: f.현재고금액, 단가: unit,
          재료비추정: e ? e.재료비 : null,
          잔량로트: e ? e.로트.join(', ') : '', 출처: 'FCO', fco: f, led: src.led(f.품목코드, opts.month), flags });
      }
    } else {
      for (const [p, e] of est) {
        const o = lots.find(x => x.제품코드 === p);
        rows.push({ 제품코드: p, 제품명: o.제품명, 규격: '', 공장: '평택', 구분: '당월 생산·미출고(로트목록 잔량)', 현재고: e.수량, 현재고금액: null,
          재료비추정: e.재료비, 잔량로트: e.로트.join(', '), 출처: '로트목록 잔량', led: src.led(p, opts.month), flags: ['FCO 칸이 비어 있어 재고액 없음(재료비만)'] });
      }
    }
    return { rows: rows.sort((a, b) => Z(b.현재고금액 ?? b.재료비추정) - Z(a.현재고금액 ?? a.재료비추정)), hasFco: !!src.Fco };
  }

  function run(book, userOpts = {}) {
    const lotInfo = readLots(book);
    const month = userOpts.month || detectMonth(book);
    if (!month) throw new Error('당월을 알 수 없습니다. 당월을 입력하세요.');
    RU = { ...DEFAULT_RULES, ...(userOpts.rules || {}) };
    const ratio = { ...RATIO, ...(userOpts.ratio || {}) };
    const opts = { month, ratio, rules: { ...RU }, history: userOpts.history || null };
    const src = prepare(book, month);
    const lots = calcLots(lotInfo.lots, src, opts);
    const prev = calcPrev(lots, src, opts);
    const taken = new Set([...lots.filter(o => o.K === 1).map(o => o.key), ...prev.map(o => o.key)]);
    // 평택비율 < 1 로트의 화성분은 lots 쪽에서 따로 계산하므로, 화성 품목에서는 K가 있는 로트 전체를 제외
    lots.filter(o => o.K != null).forEach(o => taken.add(o.key));
    const hw = calcHwaseong(taken, src, opts);
    const stock = calcStock(lots, src, opts);
    const totalShip = src.Sh.reduce((a, r) => a + r.amt, 0);
    const s = (arr, k) => arr.reduce((a, x) => a + Z(x[k]), 0);
    const P당월 = { 매출: s(lots, '매출'), 원료: s(lots, '원료'), 자재: s(lots, '자재') };
    const P전월 = { 매출: s(prev, '매출'), 원료: s(prev, '원료'), 자재: s(prev, '자재') };
    const P = { 매출: P당월.매출 + P전월.매출, 원료: P당월.원료 + P전월.원료, 자재: P당월.자재 + P전월.자재 };
    const Hl = { 매출: s(lots, 'H매출'), 원료: s(lots, 'H원료'), 자재: s(lots, 'H자재') };
    const Hu = { 매출: s(hw, '매출'), 원료: s(hw, '원료'), 자재: s(hw, '자재'), 미확인매출: hw.filter(g => g.원가 == null).reduce((a, g) => a + g.매출, 0) };
    const H = { 매출: Hl.매출 + Hu.매출, 원료: Hl.원료 + Hu.원료, 자재: Hl.자재 + Hu.자재 };
    const 미확인 = { 평택: [...lots, ...prev].filter(o => o.매출 && o.원가 == null).reduce((a, o) => a + o.매출, 0) };
    return { opts, lotSheet: lotInfo.sheet, lots, prev, hw, stock, totalShip, P, P당월, P전월, H, Hl, Hu, 미확인 };
  }
  return { run, readLots, T, N, RATIO, DEFAULT_RULES };
})();
if (typeof module !== 'undefined') module.exports = Engine;

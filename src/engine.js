// ===== 화성·평택 매출/원가 계산 엔진 (파워쿼리 로직 이식) =====
// book = { sheets: [ { name, rows: any[][] (A1 기준), formula(r,c) -> string|null } ] }
// opts = { month, fx, ratio: {생산구분: 비율} }
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
  const sum = a => { const b = a.filter(x => x != null); return b.length ? b.reduce((p, c) => p + c, 0) : null; };
  const noSpace = s => s.replace(/\s/g, '');

  function sheetBy(book, kw) {
    const s = book.sheets.find(x => noSpace(x.name).includes(kw));
    if (!s) throw new Error(`원천 파일에 '${kw}' 시트가 없습니다`);
    return s;
  }
  function promote(rows, key) {
    const lim = Math.min(rows.length, 15);
    for (let i = 0; i < lim; i++) {
      if ((rows[i] || []).some(c => Clean(c) === key)) return { hdr: rows[i].map(Clean), data: rows.slice(i + 1) };
    }
    throw new Error(`머리글을 찾지 못했습니다: ${key}`);
  }
  function pick(pr, keys) {
    const pos = keys.map(k => {
      const p = pr.hdr.indexOf(k);
      if (p < 0) throw new Error(`열을 찾지 못했습니다: ${k}`);
      return p;
    });
    return pr.data.map(r => { const o = {}; keys.forEach((k, j) => o[k] = (r || [])[pos[j]] ?? null); return o; });
  }
  const monthOf = v => {
    if (v == null) return null;
    if (v instanceof Date) return v.getMonth() + 1;
    const p = T(v).split(/[.\-/]/);
    return p.length >= 2 ? N(p[1]) : null;
  };
  const sheetMonth = name => { const m = name.match(/(\d{1,2})\s*월/); return m ? Number(m[1]) : null; };

  function leadNum(s) {
    const s1 = s.replace(/,/g, '').trim();
    let pos = -1;
    for (let i = 0; i < s1.length; i++) if (!'0123456789.'.includes(s1[i])) { pos = i; break; }
    const numTxt = pos === -1 ? s1 : s1.slice(0, pos);
    const n = numTxt === '' ? null : (isNaN(Number(numTxt)) ? null : Number(numTxt));
    return { n, rest: s1.slice(numTxt.length).trim() };
  }
  function parseSpec(spec) {
    if (!spec) return { mg: null, cap: null, ea: null };
    const parts = spec.replace(/,/g, '').split('x');
    const head = leadNum(parts[0]); const unit = head.rest.toLowerCase();
    const mg = head.n == null ? null : unit.startsWith('mg') ? head.n
      : (unit.startsWith('g') || unit.startsWith('ml')) ? head.n * 1000 : null;
    const tails = parts.slice(1).map(leadNum);
    const capL = tails.filter(t => t.n != null && ['연질캡슐', '캡슐', '포', '정'].some(p => t.rest.startsWith(p)));
    const eaL = tails.filter(t => t.n != null && ['ea', 'case', '병'].some(p => t.rest.toLowerCase().startsWith(p)));
    return { mg, cap: capL.length ? capL[0].n : 1, ea: eaL.length ? eaL[0].n : 1 };
  }

  // ---------- 로트목록 읽기 ('로트목록' 표 또는 기존 '당월 생산' 시트) ----------
  function readLots(book) {
    for (const s of book.sheets) {
      const rows = s.rows;
      const hdrIdx = [];
      rows.forEach((r, i) => {
        const c = (r || []).map(Clean);
        if (c.includes('완료수량') && c.some(x => x.startsWith('로트번호')) && c.some(x => x.startsWith('제품코드'))) hdrIdx.push(i);
      });
      if (!hdrIdx.length) continue;
      const lots = []; let fxRate = null;
      hdrIdx.forEach((hi, k) => {
        const h = rows[hi].map(Clean);
        const find = (...names) => { for (const n of names) { const p = h.indexOf(n); if (p >= 0) return p; } return -1; };
        const findStart = pre => h.findIndex(x => x.startsWith(pre));
        const cProd = find('제품코드(제공자료)') >= 0 ? find('제품코드(제공자료)') : findStart('제품코드');
        const cLot = find('로트번호(확인)', '로트번호', '로트번호(제공자료)');
        const cName = find('제품명'), cCust = find('고객사');
        const cLine = find('생산구분', '평택비율');
        const cSpec = find('규격수', '규격'), cQty = find('완료수량'), cPrice = find('완료단가');
        const cGubun = find('구분'), cFx = find('환율적용'), cAdj = find('보정계수');
        const cUnit = find('단가'), cWeight = find('생산중량(kg)', '생산중량');
        // 구간 이름: 머리글 위 3줄 안에 '전월'이 있으면 전월
        let section = '당월';
        for (let j = Math.max(0, hi - 3); j < hi; j++) if ((rows[j] || []).some(c => T(c).includes('전월'))) section = '전월';
        const end = k + 1 < hdrIdx.length ? hdrIdx[k + 1] : rows.length;
        for (let i = hi + 1; i < end; i++) {
          const r = rows[i] || [];
          const lot = T(r[cLot]); const qty = N(r[cQty]);
          if (!lot || qty == null || cProd < 0) continue;
          let fx = cFx >= 0 ? T(r[cFx]).toUpperCase() === 'Y' : false;
          if (cFx < 0 && cUnit >= 0 && s.formula) {
            const f = s.formula(i, cUnit); const m = f && f.match(/\*\s*([\d.]+)\s*$/);
            if (m) { fx = true; fxRate = Number(m[1]); }
          }
          let adj = cAdj >= 0 ? N(r[cAdj]) : null;
          if (cAdj < 0 && cWeight >= 0 && s.formula) {
            const f = s.formula(i, cWeight); const m = f && f.match(/\*\s*\$?([A-Z]{1,3})\$?(\d+)\s*$/);
            if (m) {
              const col = m[1].split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0) - 1;
              adj = N((rows[Number(m[2]) - 1] || [])[col]);
            }
          }
          lots.push({
            구분: cGubun >= 0 && T(r[cGubun]) ? T(r[cGubun]) : section,
            제품코드: T(r[cProd]), 로트번호: lot, 제품명: T(r[cName]), 고객사: T(r[cCust]),
            생산구분: T(r[cLine]), 규격수: N(r[cSpec]), 완료수량: qty, 완료단가: N(r[cPrice]),
            환율적용: fx, 보정계수: adj == null ? 1 : adj
          });
        }
      });
      if (lots.length) return { sheet: s.name, lots, fxRate };
    }
    throw new Error("로트목록을 찾지 못했습니다. '당월 생산' 시트(제공자료) 또는 '로트목록' 표가 필요합니다.");
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
    const Od = pick(promote(sheetBy(book, '발주현황').rows, 'LOT'), ['매출월', 'LOT', '제품코드', '구분', '수주수량', '수주단가'])
      .filter(r => N(r['매출월']) === month)
      .map(r => ({ ...r, lot: T(r.LOT), key: T(r.LOT) + '|' + T(r['제품코드']) }));
    const Wo = pick(promote(sheetBy(book, '작업지시').rows, '제품로트번호'), ['제품로트번호', '지시제품', '규격'])
      .map(r => ({ ...r, key: T(r['제품로트번호']) + '|' + T(r['지시제품']) }));
    const Sh = pick(promote(sheetBy(book, '출하대장').rows, '로트번호'),
      ['로트번호', '품목코드', '제품명', '매출업체', '출하량', '중량', '공급가액'])
      .filter(r => T(r['품목코드']) !== '' || N(r['공급가액']) != null)
      .map(r => ({ ...r, key: T(r['로트번호']) + '|' + T(r['품목코드']) }));
    const Mi = pick(promote(sheetBy(book, '자재투입').rows, '투입번호'), ['투입번호', '제조일자', '생산중량', '투입중량', '반제품/자재로트번호']);
    const seen = new Set(); const MiG = new Map();
    for (const r of Mi) {
      const lot = T(r['반제품/자재로트번호']); const k = lot + '\u0001' + T(r['투입번호']);
      if (seen.has(k)) continue; seen.add(k);
      if (!MiG.has(lot)) MiG.set(lot, { AJ: [], AK: [], mm: [] });
      const g = MiG.get(lot); g.AJ.push(N(r['생산중량'])); g.AK.push(N(r['투입중량'])); g.mm.push(monthOf(r['제조일자']));
    }
    for (const [k, g] of MiG) {
      const mms = g.mm.filter(x => x != null);
      MiG.set(k, { AJ: sum(g.AJ), AK: sum(g.AK), mm: mms.length ? Math.min(...mms) : null });
    }
    const Lg = new Map(); // key p|m -> {cAA,cAB,cAC,pw}
    const ledSheets = book.sheets.filter(x => noSpace(x.name).includes('단가대장'));
    if (!ledSheets.length) throw new Error("원천 파일에 '제품 단가 대장' 시트가 없습니다");
    for (const s of ledSheets) {
      const m = sheetMonth(s.name);
      for (const r of pick(promote(s.rows, '자재코드'), ['품목코드', '자재코드', '상세금액', '생산중량'])) {
        const p = T(r['품목코드']); if (!p.startsWith('1')) continue;
        const mat = T(r['자재코드']); const amt = Z(N(r['상세금액']));
        const k = p + '|' + m;
        if (!Lg.has(k)) Lg.set(k, { p, m, cAA: 0, cAB: 0, cAC: 0, pw: N(r['생산중량']) });
        const g = Lg.get(k);
        g.cAA += amt; if (mat.startsWith('3')) g.cAB += amt; if (mat === '4C00000149') g.cAC += amt;
      }
    }
    const byProd = new Map();
    for (const g of Lg.values()) { if (!byProd.has(g.p)) byProd.set(g.p, []); byProd.get(g.p).push(g); }
    return { Od, Wo, Sh, MiG, byProd };
  }

  // ---------- 평택 대상 로트 계산 ----------
  function calcLots(lots, src, opts) {
    const { Od, Wo, Sh, MiG, byProd } = src;
    const base = lots.map(l => {
      const lot = l.로트번호, prod = l.제품코드, key = lot + '|' + prod;
      const kk = opts.ratio[l.생산구분] ?? null;
      const adj = l.보정계수 == null ? 1 : l.보정계수;
      const w = Wo.filter(x => x.key === key);
      const spec = w.length ? T(w[0]['규격']) : null;
      const o1 = Od.filter(x => x.key === key);
      const o = o1.length ? o1 : Od.filter(x => x.lot === lot);
      const odFlag = o1.length ? null : (o.length ? '발주 제품코드 상이(로트로만 매칭)' : '발주 없음');
      const qty = Z(sum(o.map(x => N(x['수주수량']))));
      const p0 = o.length ? N(o[0]['수주단가']) : null;
      const price = p0 == null ? null : p0 * (l.환율적용 ? opts.fx : 1);
      const shipFull = Z(sum(Sh.filter(x => x.key === key).map(x => N(x['공급가액']))));
      const m = MiG.get(lot);
      const mm = (m && m.mm) || opts.month;
      const cand = byProd.get(prod) || [];
      const later = cand.filter(c => c.m >= mm).sort((a, b) => a.m - b.m);
      const crow = later.length ? later[0] : (cand.length ? cand.reduce((a, b) => a.m > b.m ? a : b) : null);
      const sp = parseSpec(spec);
      return {
        ...l, K: kk, Adj: adj, Acct: o.length ? T(o[0]['구분']) : '없음', OQty: qty, OPrice: price,
        ShipFull: shipFull, CM: crow ? crow.m : null,
        cAA: crow ? crow.cAA : 0, cAB: crow ? crow.cAB : 0, cAC: crow ? crow.cAC : 0,
        AJraw: m ? m.AJ : null, AK: m ? m.AK : null, Spec: spec || '확인 불가', Mg: sp.mg, Cap: sp.cap, Ea: sp.ea,
        flags: [kk == null ? '생산구분 확인' : null, w.length ? null : '작업지시 없음', odFlag,
          shipFull === 0 ? '출하 없음' : null, crow ? null : '단가대장 없음'].filter(Boolean)
      };
    });
    return base.map(r => {
      const grp = base.filter(x => x.제품코드 === r.제품코드 && x.CM === r.CM);
      const tot = grp.reduce((a, x) => a + Z(x.완료수량), 0);
      const s = tot === 0 ? 1 / grp.length : Z(r.완료수량) / tot;
      const K = r.K, a = r.Adj;
      const o = { ...r, Share: s };
      // 로트 전체(100%) 원가 → 평택분(K) / 화성분(1-K)
      o.FullAA = r.cAA * s * a; o.FullAD = (r.cAB + r.cAC) * s * a; o.FullAE = o.FullAD === 0 ? 0 : o.FullAA - o.FullAD;
      o.P = Z(r.완료수량) * Z(r.완료단가);
      o.Q = K == null ? null : o.P * K;
      o.V = (r.OPrice == null || K == null) ? null : r.OQty * r.OPrice * K;
      o.X = K == null ? null : r.ShipFull * K;
      o.AA = K == null ? null : o.FullAA * K;
      o.AB = r.cAB * s * a; o.AC = r.cAC * s * a;
      o.AD = K == null ? null : o.FullAD * K;
      o.AE = o.AD == null ? null : (o.AD === 0 ? 0 : o.AA - o.AD);
      o.AI = (o.X && o.AD != null) ? (o.AD + o.AE) / o.X : null;
      o.AJ = r.AJraw == null ? null : r.AJraw * a;
      o.AL = (r.AJraw && r.AK) ? r.AJraw / r.AK : null;
      o.AR = r.Mg == null ? null : r.Cap * r.Ea * r.Mg / 1e6;
      o.AS = (o.AJ == null || !o.AR) ? null : o.AJ / o.AR;
      o.AV = o.AS == null ? null : o.AS - r.OQty;
      o.AW = (o.AS == null || r.OQty === 0) ? null : o.AS / r.OQty;
      o.BA = (o.AW == null || o.AD == null) ? null : o.AW * o.AD;
      o.BB = (o.AW == null || o.AE == null) ? null : o.AW * o.AE;
      o.BC = (o.BA == null || !o.X) ? null : o.BA / o.X;
      o.BD = (o.BB == null || !o.X) ? null : o.BB / o.X;
      if (o.AW != null && Math.abs(o.AW - 1) > 0.01) o.flags = [...o.flags, '생산-수주 수량차이'];
      // 화성분 (화성포장 등 K<1)
      const h = K == null ? 0 : 1 - K;
      o.HX = r.ShipFull * h; o.HAD = o.FullAD * h; o.HAE = o.FullAE * h;
      return o;
    });
  }

  // ---------- 화성 품목별 (로트목록에 없는 출하) ----------
  function calcHwaseong(lotRows, src, opts) {
    const listed = new Set(lotRows.map(l => l.로트번호 + '|' + l.제품코드));
    const groups = new Map();
    for (const r of src.Sh) {
      if (listed.has(r.key)) continue;
      const p = T(r['품목코드']);
      if (!groups.has(p)) groups.set(p, { 품목코드: p, 제품명: T(r['제품명']), 거래처: new Set(), 출하량: 0, 중량: 0, 매출: 0, 건수: 0 });
      const g = groups.get(p);
      g.거래처.add(T(r['매출업체'])); g.출하량 += Z(N(r['출하량'])); g.중량 += Z(N(r['중량'])); g.매출 += Z(N(r['공급가액'])); g.건수++;
    }
    const out = [];
    for (const g of groups.values()) {
      const cand = src.byProd.get(g.품목코드) || [];
      const le = cand.filter(c => c.m <= opts.month).sort((a, b) => b.m - a.m);
      const c = le.length ? le[0] : (cand.length ? cand.reduce((a, b) => a.m > b.m ? a : b) : null);
      let unit = null, cost = null, raw = null, mat = null, flag = null;
      if (!c) flag = g.품목코드.startsWith('1') ? '단가대장 없음' : '제품 외 품목(원가 대상 아님)';
      else if (!c.pw) flag = '단가대장 생산중량 없음';
      else if (!g.중량) flag = '출하중량 없음';
      else {
        unit = c.cAA / c.pw; cost = g.중량 * unit;
        raw = c.cAA ? cost * (c.cAB + c.cAC) / c.cAA : 0; mat = cost - raw;
      }
      out.push({ ...g, 거래처: [...g.거래처].filter(Boolean).join(', '), 원가월: c ? c.m : null, kg당원가: unit,
        원료: raw, 자재: mat, 원가: cost, 원가율: (cost != null && g.매출) ? cost / g.매출 : null, flag });
    }
    return out.sort((a, b) => b.매출 - a.매출);
  }

  function run(book, userOpts = {}) {
    const lotInfo = readLots(book);
    const month = userOpts.month || detectMonth(book);
    if (!month) throw new Error('당월을 알 수 없습니다. 설정에서 당월을 입력하세요.');
    const fx = userOpts.fx || lotInfo.fxRate || 1360;
    const ratio = userOpts.ratio || { '평택': 1, '화성포장': 0.8, '멀티연질': 0.6, '정제화성': 0.2, '화성': 0 };
    const opts = { month, fx, ratio };
    const src = prepare(book, month);
    const lots = calcLots(lotInfo.lots, src, opts);
    const hw = calcHwaseong(lotInfo.lots, src, opts);
    const totalShip = src.Sh.reduce((a, r) => a + Z(N(r['공급가액'])), 0);
    const P = { 매출: sum(lots.map(o => o.X)) || 0, 원료: sum(lots.map(o => o.AD)) || 0, 자재: sum(lots.map(o => o.AE)) || 0 };
    const Hl = { 매출: sum(lots.map(o => o.HX)) || 0, 원료: sum(lots.map(o => o.HAD)) || 0, 자재: sum(lots.map(o => o.HAE)) || 0 };
    const Hu = { 매출: hw.reduce((a, g) => a + g.매출, 0), 원료: hw.reduce((a, g) => a + Z(g.원료), 0), 자재: hw.reduce((a, g) => a + Z(g.자재), 0),
      미확인매출: hw.filter(g => g.원가 == null).reduce((a, g) => a + g.매출, 0) };
    const H = { 매출: Hl.매출 + Hu.매출, 원료: Hl.원료 + Hu.원료, 자재: Hl.자재 + Hu.자재 };
    return { opts, lotSheet: lotInfo.sheet, lots, hw, totalShip, P, H, Hl, Hu };
  }
  return { run, parseSpec, readLots, T, N };
})();
if (typeof module !== 'undefined') module.exports = Engine;

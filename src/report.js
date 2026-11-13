// ===== 결과 집계·엑셀 보고서 생성 (웹과 CLI 공용) =====
const Report = (() => {
  const VERSION = 'v2.5';
  const VERSION_DATE = '2026-11-13';
  const HISTORY = [
    'v2.5 (2026-11-13) ◀ ▶로 달을 오가며 계산 결과 보기(달마다 칸·결과 기억, 보관함 달은 자동 계산), 계산 내역 칩, 보관함 열기/닫기 버튼, 보관함 칩 × 빼기',
    'v2.4 (2026-11-13) 새 버전 배포 알림, 배포마다 계산 파일 주소에 버전 번호(캐시 방지)',
    'v2.3 (2026-11-13) 보관함을 용도별 칸으로(칸마다 여러 달 파일, 저장된 달 표시), 필수 5칸 색 구분',
    'v2.2 (2026-11-13) 계산 규칙 화면에서 변경(평택비율·평택 판단·원료 코드·단가대장 월 기준·제외 단어), 보관함 자료 월 옮기기·지우기·로트목록 수정(수정 이력)',
    'v2.1 (2026-11-13) 월별 자료 보관함(미가공 파일 저장·백업), 생산월을 로트목록 생산귀속으로만 확정, 단가대장은 마감월 기준',
    'v2.0.2 (2026-11-13) 추정 계산 제거: 생산월 단가대장이 없으면 원가 빈칸, 화성은 당월 대장만, 재고 비례 추정 삭제',
    'v2.0.1 (2026-11-13) 계산·파일 읽기·엑셀 만들기 오류를 오류 창으로 표시',
    'v2.0 (2026-11-13) 로트 방식 개편(제품코드+로트, 출하분만 원가, 조정비율 제거), 전월 생산 출하, FCO 미출고 재고, 검수 화면',
    'v1.2 (2026-11-13) 자료별 칸',
    'v1.1 (2026-11-11) 시트별 개별 파일',
    'v1.0 (2026-10-30) 첫 배포',
  ];
  const div = (a, b) => (b ? a / b : null);

  function totals(R) {
    const P = R.P, H = R.H;
    return { P, H, P당월: R.P당월, P전월: R.P전월, T: { 매출: P.매출 + H.매출, 원료: P.원료 + H.원료, 자재: P.자재 + H.자재 }, Hl: R.Hl, Hu: R.Hu };
  }

  function flagsOf(R) {
    const out = [];
    R.lots.forEach(o => o.flags.forEach(f => out.push({ 구분: '평택 당월 로트', 코드: o.제품코드, 대상: o.로트번호, 이름: o.제품명, 내용: f, 금액: o.shipAmt })));
    R.prev.forEach(o => o.flags.forEach(f => out.push({ 구분: '평택 전월 생산', 코드: o.제품코드, 대상: o.로트번호, 이름: o.제품명, 내용: f, 금액: o.매출 })));
    R.hw.filter(g => g.flag).forEach(g => out.push({ 구분: '화성 품목', 코드: g.품목코드, 대상: '', 이름: g.제품명, 내용: g.flag, 금액: g.매출 }));
    R.stock.rows.forEach(r => r.flags.forEach(f => out.push({ 구분: '미출고 재고', 코드: r.제품코드, 대상: r.공장, 이름: r.제품명, 내용: f, 금액: r.현재고금액 })));
    return out;
  }

  function sheetFrom(XLSX, aoa, formats, widths, hdrRow = 0) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const rg = XLSX.utils.decode_range(ws['!ref']);
    for (let r = hdrRow + 1; r <= rg.e.r; r++) for (let c = 0; c <= rg.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.t === 'n' && formats[c]) cell.z = formats[c];
    }
    ws['!cols'] = widths.map(w => ({ wch: w }));
    return ws;
  }

  function buildWorkbook(XLSX, R, sourceName) {
    const S = totals(R), M = '#,##0', P1 = '0.0%', D = '#,##0.00';
    const wb = XLSX.utils.book_new();
    const line = (n, x) => [n, x.매출, x.원료, x.자재, x.원료 + x.자재, div(x.원료 + x.자재, x.매출), div(x.매출, R.totalShip)];
    const pt = R.stock.rows.filter(r => r.공장 === '평택');
    const sum = [
      ['구분', '매출', '원료', '자재', '재료원가', '재료원가율', '매출 비중'],
      line('평택', S.P), line('  평택 당월 생산', S.P당월), line('  평택 전월 이전 생산', S.P전월),
      line('화성', S.H), line('  화성 출하 품목', S.Hu), line('  평택 로트의 화성분', S.Hl), line('합계', S.T),
      [], ['기준월', R.opts.month + '월'], ['원천 파일', sourceName || ''], ['출하대장 공급가액 합계', R.totalShip],
      ['평택 당월 로트 수', R.lots.length], ['평택 전월 생산 출하 로트 수', R.prev.length], ['화성 품목 수', R.hw.length],
      ['원가 미확인 매출(평택)', R.미확인.평택], ['원가 미확인 매출(화성)', R.Hu.미확인매출],
      ['평택 미출고 기말 제품재고', R.stock.hasFco ? pt.reduce((a, r) => a + (r.현재고금액 || 0), 0) : '(FCO 칸을 넣으면 계산)'],
      [], ['적용 규칙'],
      ['평택비율', Object.entries(R.opts.ratio).map(([k, v]) => `${k} ${Math.round(v * 100)}%`).join(', ')],
      ['평택 판단(로트목록 밖 출하)', R.opts.rules ? `작업지시번호 ${R.opts.rules.ptWo || '(없음)'}… 또는 현장 ${R.opts.rules.ptSite || '(없음)'}` : ''],
      ['원료 코드', R.opts.rules ? `반제품 ${R.opts.rules.semiPrefix || '(없음)'}… + 주정 ${(R.opts.rules.alcohol || []).join(', ') || '(없음)'}` : ''],
      [], ['계산 기준'],
      ['평택 매출 = 로트목록 로트의 출하금액(제품코드+로트번호) × 평택비율(생산공장·포장공장: 평택/평택 100%, 평택/화성 80%)'],
      ['평택 당월 원가 = 출하량 × (당월 단가대장 제조금액 ÷ 같은 제품 로트들의 생산수량 합) × 평택비율. 출하 안 된 수량은 원가에 넣지 않음'],
      ['평택 전월 생산 = 로트목록에 없는 평택 출하(작업지시 10-30 또는 비오팜2공장), 원가 = 출하중량 × 생산월 단가대장 kg당 제조금액. 생산월은 이전 달 로트목록 생산귀속으로만, 없으면 빈칸'],
      ['화성 = 나머지 출하, 원가 = 출하중량 × (당월 단가대장 제조금액 ÷ 생산중량), 당월 대장에 없으면 빈칸. 원료 = 반제품(3…)+주정, 자재 = 나머지'],
      ['다른 달 자료로 추정하지 않음: 근거 자료가 없는 원가는 빈칸으로 두고 확인사항에 표시'],
      ['단가대장 금액은 재료비 기준. FCO 금액은 가공비 포함 원가로 보임'],
      [], ['계산기 버전', `${VERSION} · ${VERSION_DATE}`], ['변경 이력'], ...HISTORY.map(h => ['', h]),
    ];
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, sum, { 1: M, 2: M, 3: M, 4: M, 5: P1, 6: P1 }, [30, 16, 16, 16, 16, 11, 10]), '요약');

    const lh = ['No', '제품코드', '로트번호', '제품명', '고객사', '생산구분', '평택비율', '완료수량', '완료단가', '출하량', '출하중량', '출하금액',
      '평택 매출', '원가 방식', '원가월', '생산수량', '1개당 원가', '평택 재료원가', '원료', '자재', '재료원가율', '잔량', '잔량 재료비',
      '수주수량', '수주단가', '단가대장 생산중량', '로트 중량 합', '화성분 매출', '화성분 원료', '화성분 자재', '확인사항'];
    const lr = R.lots.map((o, i) => [i + 1, o.제품코드, o.로트번호, o.제품명, o.고객사, o.생산구분, o.K, o.완료수량, o.완료단가, o.shipQty, o.shipW, o.shipAmt,
      o.매출, o.원가방식, o.원가월, o.생산수량, o.개당원가, o.원가, o.원료, o.자재, o.원가율, o.잔량, o.잔량재료비,
      o.수주수량, o.수주단가, o.단가대장중량, o.로트중량합, o.H매출, o.H원료, o.H자재, o.flags.join(', ')]);
    const lf = {}; [7, 8, 9, 11, 12, 15, 17, 18, 19, 21, 22, 23, 24, 27, 28, 29].forEach(c => (lf[c] = M)); [6, 20].forEach(c => (lf[c] = P1)); [10, 16, 25, 26].forEach(c => (lf[c] = D));
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [lh, ...lr], lf, lh.map((h, i) => ([3, 4].includes(i) ? 30 : i === 30 ? 40 : i === 13 ? 22 : Math.max(9, h.length * 1.8)))), '평택_당월로트');

    const ph = ['No', '제품코드', '로트번호', '제품명', '거래처', '평택 근거', '작업지시일', '출하량', '출하중량', '매출', '원가월', '재료원가', '원료', '자재', '재료원가율', '확인사항'];
    const pr = R.prev.map((o, i) => [i + 1, o.제품코드, o.로트번호, o.제품명, o.고객사, o.근거, o.지시일자, o.shipQty, o.shipW, o.매출, o.원가월, o.원가, o.원료, o.자재, o.원가율, o.flags.join(', ')]);
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [ph, ...pr], { 7: M, 8: D, 9: M, 11: M, 12: M, 13: M, 14: P1 }, [5, 12, 11, 30, 26, 22, 12, 12, 11, 14, 7, 14, 14, 14, 10, 36]), '평택_전월생산');

    const hh = ['구분', '품목코드', '로트번호', '제품명', '거래처', '출하건수', '출하량', '출하중량(kg)', '매출', '원가월', 'kg당 원가', '원료', '자재', '재료원가', '재료원가율', '확인사항'];
    const hr = [
      ...R.hw.map(g => ['화성 출하', g.품목코드, '', g.제품명, g.거래처, g.건수, g.출하량, g.중량, g.매출, g.원가월, g.kg당원가, g.원료, g.자재, g.원가, g.원가율, g.flag || '']),
      ...R.lots.filter(o => o.H매출).map(o => ['평택 로트의 화성분', o.제품코드, o.로트번호, o.제품명, o.고객사, null, null, null, o.H매출, o.원가월, null,
        o.H원료, o.H자재, o.H원료 + o.H자재, div(o.H원료 + o.H자재, o.H매출), `${o.생산구분} ${Math.round((1 - o.K) * 100)}%`]),
    ];
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [hh, ...hr], { 6: M, 7: D, 8: M, 10: M, 11: M, 12: M, 13: M, 14: P1 },
      [16, 12, 12, 30, 28, 8, 12, 12, 15, 7, 11, 15, 15, 15, 10, 26]), '화성_품목별');

    const sh = ['공장', '제품코드', '제품명', '규격', '구분', '전재고', '입고', '출고', '현재고', '현재고금액(FCO)', '단가', '재료비(로트 잔량)', '당월 잔량 로트', '출처', '확인사항'];
    const sr = R.stock.rows.map(r => [r.공장, r.제품코드, r.제품명, r.규격, r.구분, r.전재고, r.입고, r.출고, r.현재고, r.현재고금액, r.단가, r.재료비추정, r.잔량로트, r.출처, r.flags.join(', ')]);
    const tot = k => R.stock.rows.reduce((a, r) => a + (r[k] || 0), 0);
    sr.push([], ['평택 합계', '', '', '', '', null, null, null, null, pt.reduce((a, r) => a + (r.현재고금액 || 0), 0), null, pt.reduce((a, r) => a + (r.재료비추정 || 0), 0)],
      ['전체 합계', '', '', '', '', null, null, null, null, tot('현재고금액'), null, tot('재료비추정')]);
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [sh, ...sr], { 5: M, 6: M, 7: M, 8: M, 9: M, 10: M, 11: M }, [8, 12, 32, 18, 20, 12, 12, 12, 12, 15, 11, 15, 26, 12, 30]), '미출고_재고');

    const fl = flagsOf(R);
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [['구분', '코드', '로트·공장', '제품명', '내용', '금액'],
      ...fl.map(f => [f.구분, f.코드, f.대상, f.이름, f.내용, f.금액])], { 5: M }, [14, 12, 12, 32, 44, 15]), '확인사항');
    return wb;
  }

  const outputName = R => `화성평택_매출원가_${R.opts.month}월.xlsx`;
  return { totals, flagsOf, buildWorkbook, outputName, VERSION, VERSION_DATE, HISTORY };
})();
if (typeof module !== 'undefined') module.exports = Report;

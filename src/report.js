// ===== 결과 집계·엑셀 보고서 생성 (웹과 CLI 공용) =====
const Report = (() => {
  const div = (a, b) => (b ? a / b : null);
  const pct0 = v => (v == null ? '' : Math.round(v * 100) + '%');

  function totals(R) {
    const P = R.P, H = R.H;
    return { P, H, T: { 매출: P.매출 + H.매출, 원료: P.원료 + H.원료, 자재: P.자재 + H.자재 }, Hl: R.Hl, Hu: R.Hu };
  }

  function flagsOf(R) {
    const out = [];
    R.lots.forEach(o => o.flags.forEach(f => out.push({ 구분: '평택 로트', 코드: o.제품코드, 대상: o.로트번호, 이름: o.제품명, 내용: f, 금액: o.X })));
    R.hw.filter(g => g.flag).forEach(g => out.push({ 구분: '화성 품목', 코드: g.품목코드, 대상: '', 이름: g.제품명, 내용: g.flag, 금액: g.매출 }));
    return out;
  }

  function sheetFrom(XLSX, aoa, formats, widths) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const rg = XLSX.utils.decode_range(ws['!ref']);
    for (let r = 1; r <= rg.e.r; r++) for (let c = 0; c <= rg.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.t === 'n' && formats[c]) cell.z = formats[c];
    }
    ws['!cols'] = widths.map(w => ({ wch: w }));
    return ws;
  }

  // R = Engine.run() 결과, sourceName = 원천 파일 이름
  function buildWorkbook(XLSX, R, sourceName) {
    const S = totals(R), M = '#,##0', P1 = '0.0%', D = '#,##0.00';
    const wb = XLSX.utils.book_new();
    const line = (n, x) => [n, x.매출, x.원료, x.자재, x.원료 + x.자재, div(x.원료, x.매출), div(x.자재, x.매출), div(x.원료 + x.자재, x.매출), div(x.매출, R.totalShip)];
    const ratioText = Object.entries(R.opts.ratio).map(([k, v]) => `${k} ${pct0(v)}`).join(', ');
    const sum = [
      ['구분', '매출', '원료', '자재', '재료원가', '원료율', '자재율', '재료원가율', '매출 비중'],
      line('평택', S.P), line('화성', S.H), line('  화성 출하 품목', S.Hu), line('  평택 로트의 화성분', S.Hl), line('합계', S.T),
      [], ['기준월', R.opts.month + '월'], ['원천 파일', sourceName || ''], ['출하대장 공급가액 합계', R.totalShip],
      ['평택 로트 수', R.lots.length], ['화성 품목 수', R.hw.length], ['원가 미확인 화성 매출', R.Hu.미확인매출], ['환율', R.opts.fx],
      [], ['계산 기준'],
      [`평택 매출 = 로트목록 로트의 출하금액 × 평택비율(${ratioText})`],
      ['화성 매출 = 출하대장 전체 − 평택 매출 (로트목록에 없는 출하 + 평택 로트의 화성분)'],
      ['평택 원가 = 로트 원가월 단가대장 제품원가 × 완료수량 배분비율 × 평택비율 × 보정계수'],
      ['화성 원가 = 출하중량 × (단가대장 제품 제조금액 ÷ 생산중량), 원료 = 반제품+주정, 자재 = 나머지'],
    ];
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, sum, { 1: M, 2: M, 3: M, 4: M, 5: P1, 6: P1, 7: P1, 8: P1 }, [26, 16, 16, 16, 16, 9, 9, 11, 10]), '요약');

    const ph = ['구분', 'No', '제품코드', '로트번호', '제품명', '고객사', '생산구분', '평택비율', '규격(제공)', '완료수량', '완료단가', '완료금액', '완료금액(비율)',
      '매출계정', '수주수량', '수주단가', '수주매출액', '출하금액(평택분)', '원가월', '배분비율', '제품금액', '반제품금액', '주정', '원료금액', '자재금액', '원부자재%',
      '생산중량(kg)', '투입중량(kg)', '수율', '규격', '규격mg', '캡슐/포', 'set ea', '1ea중량(kg)', '생산중량→set', '수량차이', '조정비율', '원료원가', '부자재원가', '원료%', '부자재%',
      '환율적용', '보정계수', '출하금액(로트전체)', '화성분 매출', '화성분 원료', '화성분 자재', '확인사항'];
    const pr = R.lots.map((o, i) => [o.구분, i + 1, o.제품코드, o.로트번호, o.제품명, o.고객사, o.생산구분, o.K, o.규격수, o.완료수량, o.완료단가, o.P, o.Q,
      o.Acct, o.OQty, o.OPrice, o.V, o.X, o.CM, o.Share, o.AA, o.AB, o.AC, o.AD, o.AE, o.AI,
      o.AJ, o.AK, o.AL, o.Spec, o.Mg, o.Cap, o.Ea, o.AR, o.AS, o.AV, o.AW, o.BA, o.BB, o.BC, o.BD,
      o.환율적용 ? 'Y' : '', o.보정계수, o.ShipFull, o.HX, o.HAD, o.HAE, o.flags.join(', ')]);
    const pf = {};
    [9, 10, 11, 12, 14, 15, 16, 17, 20, 21, 22, 23, 24, 37, 38, 43, 44, 45, 46].forEach(c => (pf[c] = M));
    [7, 19, 25, 28, 39, 40].forEach(c => (pf[c] = P1));
    [26, 27, 34, 35].forEach(c => (pf[c] = D));
    pf[36] = '0.0000'; pf[42] = '0.0000'; pf[33] = '0.000000';
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [ph, ...pr], pf,
      ph.map((h, i) => ([4, 5].includes(i) ? 30 : i === 29 ? 26 : i === 47 ? 30 : Math.max(9, h.length * 1.6)))), '평택_로트별');

    const hh = ['구분', '품목코드', '로트번호', '제품명', '거래처', '출하건수', '출하량', '출하중량(kg)', '매출', '원가월', 'kg당 원가', '원료', '자재', '재료원가', '재료원가율', '확인사항'];
    const hr = [
      ...R.hw.map(g => ['화성 출하', g.품목코드, '', g.제품명, g.거래처, g.건수, g.출하량, g.중량, g.매출, g.원가월, g.kg당원가, g.원료, g.자재, g.원가, g.원가율, g.flag || '']),
      ...R.lots.filter(o => o.HX).map(o => ['평택 로트의 화성분', o.제품코드, o.로트번호, o.제품명, o.고객사, null, null, null, o.HX, o.CM, null,
        o.HAD, o.HAE, o.HAD + o.HAE, div(o.HAD + o.HAE, o.HX), `${o.생산구분} ${pct0(1 - o.K)}`]),
    ];
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [hh, ...hr], { 6: M, 7: D, 8: M, 10: M, 11: M, 12: M, 13: M, 14: P1 },
      [16, 12, 12, 30, 28, 8, 12, 12, 15, 7, 11, 15, 15, 15, 10, 26]), '화성_품목별');

    const fl = flagsOf(R);
    XLSX.utils.book_append_sheet(wb, sheetFrom(XLSX, [['구분', '코드', '로트', '제품명', '내용', '금액'],
      ...fl.map(f => [f.구분, f.코드, f.대상, f.이름, f.내용, f.금액])], { 5: M }, [10, 12, 12, 32, 32, 15]), '확인사항');
    return wb;
  }

  const outputName = R => `화성평택_매출원가_${R.opts.month}월.xlsx`;
  return { totals, flagsOf, buildWorkbook, outputName };
})();
if (typeof module !== 'undefined') module.exports = Report;

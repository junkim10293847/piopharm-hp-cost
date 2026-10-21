// SheetJS 통합문서 → 엔진 입력 (A1 기준 2차원 배열 + 수식 조회)
function bookFromWorkbook(XLSX, wb) {
  return {
    sheets: wb.SheetNames.map(name => {
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
    })
  };
}
if (typeof module !== 'undefined') module.exports = bookFromWorkbook;

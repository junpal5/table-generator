/*
 * 계산된 표를 서식이 들어간 엑셀(.xlsx) 파일로 만듭니다. (ExcelJS 사용)
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  const COLOR = {
    head: 'FFDCE6F2',
    group: 'FFF2F5FA',
    total: 'FFFFF7E0',
    stat: 'FFF3F3F3',
    border: 'FFB4BCC8',
  };
  const thin = { style: 'thin', color: { argb: COLOR.border } };
  const BORDER = { top: thin, left: thin, bottom: thin, right: thin };

  function numFmt(fmt, opts) {
    const d = TG.numDigits(fmt, opts.decimals == null ? 1 : opts.decimals);
    return d === 0 ? '#,##0' : '#,##0.' + '0'.repeat(d);
  }

  function sheetName(s, used) {
    let base = String(s).replace(/[\\\/\?\*\[\]:]/g, ' ').trim().slice(0, 28) || '표';
    let name = base;
    let i = 2;
    while (used.has(name)) name = base.slice(0, 26) + '_' + i++;
    used.add(name);
    return name;
  }

  // 표 하나를 시트의 startRow부터 씁니다. 다음 빈 행 번호를 돌려줍니다.
  function writeTable(ws, table, opts, startRow) {
    let r = startRow;
    const title = ws.getCell(r, 1);
    title.value = `표 ${table.no}. ${table.title}`;
    title.font = { bold: true, size: 11 };
    r++;
    const vars = ws.getCell(r, 1);
    vars.value = `변수: ${table.vars.join(', ')}`;
    vars.font = { size: 9, color: { argb: 'FF7F7F7F' } };
    r++;
    if (table.kind === 'error') {
      ws.getCell(r, 1).value = table.notes.join(' ');
      return r + 3;
    }
    const grid = TG.toGrid(table, opts);
    const headRows = (opts.orientation || 'row') === 'row' ? 1 : 2;
    grid.forEach((line, i) => {
      line.forEach((c, j) => {
        const cell = ws.getCell(r + i, j + 1);
        cell.border = BORDER;
        if (c.merged) return;
        if (c.t === 'n') {
          cell.value = c.v == null || !isFinite(c.v) ? '-' : c.v;
          cell.numFmt = numFmt(c.fmt, opts);
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else {
          cell.value = c.v;
          cell.alignment = { horizontal: c.t === 'h' ? 'center' : 'left', vertical: 'middle', wrapText: true };
        }
        let fill = null;
        if (i < headRows || c.t === 'h') fill = COLOR.head;
        else if (c.total) fill = COLOR.total;
        else if (c.t === 'g') fill = COLOR.group;
        else if (c.stat) fill = COLOR.stat;
        if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        if (i < headRows || c.t === 'g' || c.total) cell.font = { bold: true, size: 10 };
        else cell.font = { size: 10 };
        if (c.rs || c.cs) ws.mergeCells(r + i, j + 1, r + i + (c.rs || 1) - 1, j + 1 + (c.cs || 1) - 1);
      });
    });
    r += grid.length;
    table.notes.forEach((n) => {
      const cell = ws.getCell(r, 1);
      cell.value = '※ ' + n;
      cell.font = { size: 9, color: { argb: 'FF595959' } };
      r++;
    });
    return r + 2;
  }

  function setWidths(ws, opts) {
    const row = (opts.orientation || 'row') === 'row';
    ws.getColumn(1).width = row ? 14 : 26;
    ws.getColumn(2).width = row ? 18 : 11;
    for (let c = 3; c <= 60; c++) ws.getColumn(c).width = 11;
  }

  async function buildWorkbook(result, opts) {
    const ExcelJS = root.ExcelJS;
    const wb = new ExcelJS.Workbook();
    wb.creator = '설문 통계표 생성기';
    const used = new Set();
    const toc = wb.addWorksheet(sheetName('목차', used));
    toc.columns = [{ width: 8 }, { width: 70 }, { width: 28 }];
    toc.addRow(['설문 통계표']).font = { bold: true, size: 14 };
    const info = [
      `생성일: ${new Date().toLocaleString('ko-KR')}`,
      `전체 사례수: ${result.n.toLocaleString('ko-KR')}명`,
      `배너: ${opts.banners && opts.banners.length ? opts.banners.map((b) => b.label).join(', ') : '없음(전체만)'}`,
      `가중치: ${opts.weightVar ? opts.weightVar + ' 적용' : '적용 안 함'}`,
    ];
    info.forEach((t) => (toc.addRow([t]).font = { size: 10, color: { argb: 'FF595959' } }));
    toc.addRow([]);
    const hr = toc.addRow(['번호', '표 제목', '변수']);
    hr.eachCell((c) => {
      c.font = { bold: true };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: COLOR.head } };
      c.border = BORDER;
    });

    const locations = [];
    if (opts.sheetMode === 'each') {
      result.tables.forEach((t) => {
        const ws = wb.addWorksheet(sheetName(`표${t.no}`, used));
        setWidths(ws, opts);
        writeTable(ws, t, opts, 1);
        locations.push({ sheet: ws.name, row: 1 });
      });
    } else {
      const ws = wb.addWorksheet(sheetName('통계표', used));
      setWidths(ws, opts);
      let r = 1;
      result.tables.forEach((t) => {
        locations.push({ sheet: ws.name, row: r });
        r = writeTable(ws, t, opts, r);
      });
    }

    result.tables.forEach((t, i) => {
      const loc = locations[i];
      const row = toc.addRow([t.no, { text: t.title, hyperlink: `#'${loc.sheet}'!A${loc.row}` }, t.vars.join(', ')]);
      row.getCell(2).font = { color: { argb: 'FF1F5FBF' }, underline: true };
      row.eachCell((c) => (c.border = BORDER));
    });
    return wb;
  }

  async function downloadExcel(result, opts, filename) {
    const wb = await buildWorkbook(result, opts);
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename || '통계표.xlsx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  Object.assign(TG, { buildWorkbook, downloadExcel });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

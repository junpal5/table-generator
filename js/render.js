/*
 * 계산된 표를 격자(grid)로 바꾸고, 화면용 HTML을 만듭니다.
 * 같은 격자를 엑셀 저장(export.js)에서도 사용합니다.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  function fmtNumber(v, fmt, opts) {
    if (v == null || !isFinite(v)) return '-';
    const d = opts.decimals == null ? 1 : opts.decimals;
    const digits = numDigits(fmt, d);
    return v.toLocaleString('ko-KR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
  }

  function numDigits(fmt, d) {
    if (fmt === 'int' || fmt === 'wn') return 0;
    if (fmt === 'mean') return d + 1;
    return d; // pct, mean1
  }

  function baseCols(opts) {
    const cols = [];
    const weighted = !!opts.weightVar;
    const mode = weighted ? opts.baseMode || 'raw' : 'raw';
    if (mode === 'raw' || mode === 'both') cols.push({ label: weighted ? '사례수(가중 전)' : '사례수', key: 'n', fmt: 'int' });
    if (mode === 'weighted' || mode === 'both') cols.push({ label: '사례수(가중 후)', key: 'wn', fmt: 'wn' });
    return cols;
  }

  function visibleColumns(table, opts) {
    return table.columns.filter((c) => !(c.isSum && opts.showSum === false));
  }

  function rowValues(table, row, opts) {
    return table.columns.map((c, i) => [c, row.values[i]]).filter(([c]) => !(c.isSum && opts.showSum === false)).map(([, v]) => v);
  }

  /**
   * 표 → 격자
   *  cell = { v, t, fmt, merged, rs, cs, total, stat }
   *   t: 'h' 머리글, 'g' 배너 그룹, 'l' 행 레이블, 'n' 숫자
   */
  function toGrid(table, opts) {
    const cols = visibleColumns(table, opts);
    const bcols = baseCols(opts);
    const grid = [];
    const cell = (v, t, extra) => Object.assign({ v, t }, extra || {});
    const skip = () => ({ merged: true });

    if ((opts.orientation || 'row') === 'row') {
      // 행: 전체/배너, 열: 사례수 + 보기
      const head = [cell('구분', 'h', { cs: 2 }), skip()];
      bcols.forEach((b) => head.push(cell(b.label, 'h')));
      cols.forEach((c) => head.push(cell(c.label, 'h', { stat: c.isStat || c.isSum })));
      grid.push(head);
      const rows = table.rows;
      for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const line = [];
        if (r.isTotal) {
          line.push(cell('전체', 'g', { cs: 2, total: true }), skip());
        } else {
          const first = i === 0 || rows[i - 1].group !== r.group || rows[i - 1].isTotal;
          if (first) {
            let span = 1;
            while (i + span < rows.length && rows[i + span].group === r.group && !rows[i + span].isTotal) span++;
            line.push(cell(r.group, 'g', { rs: span }));
          } else {
            line.push(skip());
          }
          line.push(cell(r.label, 'l'));
        }
        bcols.forEach((b) => line.push(cell(r[b.key], 'n', { fmt: b.fmt, total: r.isTotal, base: true })));
        rowValues(table, r, opts).forEach((v, j) => line.push(cell(v, 'n', { fmt: cols[j].fmt, total: r.isTotal, stat: cols[j].isStat || cols[j].isSum })));
        grid.push(line);
      }
    } else {
      // 열: 전체/배너, 행: 사례수 + 보기
      const rows = table.rows;
      const h1 = [cell('구분', 'h', { rs: 2 }), cell('전체', 'h', { rs: 2, total: true })];
      const h2 = [skip(), skip()];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (rows[i - 1].group !== r.group || rows[i - 1].isTotal) {
          let span = 1;
          while (i + span < rows.length && rows[i + span].group === r.group) span++;
          h1.push(cell(r.group, 'h', { cs: span }));
        } else {
          h1.push(skip());
        }
        h2.push(cell(r.label, 'h'));
      }
      grid.push(h1, h2);
      bcols.forEach((b) => {
        grid.push([cell(b.label, 'l', { base: true })].concat(rows.map((r) => cell(r[b.key], 'n', { fmt: b.fmt, total: r.isTotal, base: true }))));
      });
      cols.forEach((c, j) => {
        grid.push([cell(c.label, 'l', { stat: c.isStat || c.isSum })].concat(rows.map((r) => cell(rowValues(table, r, opts)[j], 'n', { fmt: c.fmt, total: r.isTotal, stat: c.isStat || c.isSum }))));
      });
    }
    return grid;
  }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
  }

  function tableHTML(table, opts) {
    if (table.kind === 'error') {
      return `<section class="tbl" id="t${table.no}"><h3>표 ${table.no}. ${esc(table.title)}</h3><p class="err">${esc(table.notes.join(' '))}</p></section>`;
    }
    const grid = toGrid(table, opts);
    const headRows = (opts.orientation || 'row') === 'row' ? 1 : 2;
    let html = `<section class="tbl" id="t${table.no}"><h3><span class="no">표 ${table.no}.</span> ${esc(table.title)}</h3>`;
    html += `<div class="vars">${esc(table.vars.join(', '))}</div><div class="scroll"><table>`;
    grid.forEach((line, r) => {
      html += r < headRows ? '<tr class="head">' : '<tr>';
      line.forEach((c) => {
        if (c.merged) return;
        const tag = c.t === 'h' ? 'th' : c.t === 'n' ? 'td' : 'th';
        const cls = [c.t, c.total ? 'total' : '', c.stat ? 'stat' : '', c.base ? 'base' : ''].filter(Boolean).join(' ');
        const span = (c.rs ? ` rowspan="${c.rs}"` : '') + (c.cs ? ` colspan="${c.cs}"` : '');
        const text = c.t === 'n' ? fmtNumber(c.v, c.fmt, opts) : esc(c.v);
        html += `<${tag} class="${cls}"${span}>${text}</${tag}>`;
      });
      html += '</tr>';
    });
    html += '</table></div>';
    if (table.notes.length) html += `<ul class="notes">${table.notes.map((n) => `<li>${esc(n)}</li>`).join('')}</ul>`;
    return html + '</section>';
  }

  Object.assign(TG, { toGrid, tableHTML, fmtNumber, numDigits, esc });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

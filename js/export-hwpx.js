/*
 * 계산된 표를 한글(HWPX) 문서로 만듭니다.
 * 빈 문서 틀(js/hwpx-template.js)에 표 서식(테두리·배경·글자·문단 모양)을 더하고,
 * 본문(section0.xml)에 표 제목 · 표 · 주석을 차례로 넣은 뒤 JSZip으로 묶습니다.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  // HWPUNIT: 1mm ≈ 283.46, 1pt = 100
  // A4 세로(210×297mm). 표가 한 쪽에 최대한 들어가도록 여백을 줄임
  //   좌우 15mm, 위아래 15mm, 머리말·꼬리말 영역 10mm
  const PAGE = { width: 59528, height: 84186, left: 4252, right: 4252, top: 4252, bottom: 4252, header: 2835, footer: 2835 };
  const TEXT_WIDTH = PAGE.width - PAGE.left - PAGE.right; // 본문 폭 180mm
  const TEXT_HEIGHT = PAGE.height - PAGE.top - PAGE.bottom - PAGE.header - PAGE.footer; // 본문 높이 약 247mm
  const CELL_MARGIN = { left: 283, right: 283, top: 85, bottom: 85 };
  const LINE = 1.3; // 줄 간격 130%
  // 표가 길면 글자를 한 단계씩 줄여 한 쪽에 들어가게 함 (가장 작은 크기로도 넘치면 다음 쪽으로 이어짐)
  const SIZES = [8.5, 8, 7.5, 7, 6.5, 6];
  const TITLE_SIZE = 10.5;
  const NOTE_SIZE = 8;

  const COLOR = {
    border: '#8C96A5',
    head: '#DCE6F2',
    group: '#F2F5FA',
    total: '#FFF7E0',
    stat: '#F3F3F3',
  };

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  }

  // ------------------------------------------------------------------
  // header.xml: 표에 쓸 서식 추가
  // ------------------------------------------------------------------
  function addToList(xml, listTag, itemsXml, count) {
    const re = new RegExp(`<hh:${listTag} itemCnt="(\\d+)">`);
    const m = re.exec(xml);
    if (!m) throw new Error(`한글 틀에서 ${listTag}를 찾지 못했습니다.`);
    const n = Number(m[1]);
    xml = xml.replace(re, `<hh:${listTag} itemCnt="${n + count}">`);
    return xml.replace(`</hh:${listTag}>`, itemsXml + `</hh:${listTag}>`);
  }

  function maxId(xml, tag) {
    let max = -1;
    const re = new RegExp(`<hh:${tag} [^>]*?\\bid="(\\d+)"`, 'g');
    let m;
    while ((m = re.exec(xml))) max = Math.max(max, Number(m[1]));
    return max;
  }

  function borderFillXml(id, fill) {
    const b = (side) => `<hh:${side} type="SOLID" width="0.12 mm" color="${COLOR.border}"/>`;
    const brush = fill ? `<hc:fillBrush><hc:winBrush faceColor="${fill}" hatchColor="#FF000000" alpha="0"/></hc:fillBrush>` : '';
    return (
      `<hh:borderFill id="${id}" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">` +
      '<hh:slash type="NONE" Crooked="0" isCounter="0"/><hh:backSlash type="NONE" Crooked="0" isCounter="0"/>' +
      b('leftBorder') + b('rightBorder') + b('topBorder') + b('bottomBorder') +
      `<hh:diagonal type="NONE" width="0.12 mm" color="${COLOR.border}"/>` +
      brush +
      '</hh:borderFill>'
    );
  }

  function charPrXml(id, { size, bold, color }) {
    const seven = (v) => `hangul="${v}" latin="${v}" hanja="${v}" japanese="${v}" other="${v}" symbol="${v}" user="${v}"`;
    return (
      `<hh:charPr id="${id}" height="${Math.round(size * 100)}" textColor="${color || '#000000'}" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="2">` +
      `<hh:fontRef ${seven(0)}/><hh:ratio ${seven(100)}/><hh:spacing ${seven(0)}/><hh:relSz ${seven(100)}/><hh:offset ${seven(0)}/>` +
      (bold ? '<hh:bold/>' : '') +
      '<hh:underline type="NONE" shape="SOLID" color="#000000"/><hh:strikeout shape="NONE" color="#000000"/><hh:outline type="NONE"/>' +
      '<hh:shadow type="NONE" color="#C0C0C0" offsetX="10" offsetY="10"/></hh:charPr>'
    );
  }

  function paraPrXml(id, align, next, keepWithNext) {
    const margin = `<hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="${next || 0}" unit="HWPUNIT"/></hh:margin><hh:lineSpacing type="PERCENT" value="130" unit="HWPUNIT"/>`;
    return (
      `<hh:paraPr id="${id}" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="0" suppressLineNumbers="0" checked="0" textDir="LTR">` +
      `<hh:align horizontal="${align}" vertical="CENTER"/><hh:heading type="NONE" idRef="0" level="0"/>` +
      `<hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="BREAK_WORD" widowOrphan="0" keepWithNext="${keepWithNext ? 1 : 0}" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>` +
      '<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>' +
      `<hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">${margin}</hp:case><hp:default>${margin}</hp:default></hp:switch>` +
      '<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0" offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/></hh:paraPr>'
    );
  }

  function buildHeader(xml) {
    const bf0 = maxId(xml, 'borderFill') + 1;
    const cp0 = maxId(xml, 'charPr') + 1;
    const pp0 = maxId(xml, 'paraPr') + 1;
    const ids = {
      bf: { plain: bf0, head: bf0 + 1, group: bf0 + 2, total: bf0 + 3, stat: bf0 + 4 },
      cp: { title: cp0, note: cp0 + 1, cover: cp0 + 2 },
      // 글자 크기별 표 글자 모양: sizes[i] = { cell, bold }
      sizes: SIZES.map((size, i) => ({ size, cell: cp0 + 3 + i * 2, bold: cp0 + 4 + i * 2 })),
      pp: { center: pp0, right: pp0 + 1, left: pp0 + 2, title: pp0 + 3 },
    };
    xml = addToList(
      xml,
      'borderFills',
      borderFillXml(ids.bf.plain) + borderFillXml(ids.bf.head, COLOR.head) + borderFillXml(ids.bf.group, COLOR.group) +
        borderFillXml(ids.bf.total, COLOR.total) + borderFillXml(ids.bf.stat, COLOR.stat),
      5,
    );
    xml = addToList(
      xml,
      'charProperties',
      charPrXml(ids.cp.title, { size: TITLE_SIZE, bold: true }) + charPrXml(ids.cp.note, { size: NOTE_SIZE, color: '#595959' }) + charPrXml(ids.cp.cover, { size: 16, bold: true }) +
        ids.sizes.map((z) => charPrXml(z.cell, { size: z.size }) + charPrXml(z.bold, { size: z.size, bold: true })).join(''),
      3 + ids.sizes.length * 2,
    );
    // 표 제목은 다음 문단(표)과 같은 쪽에 두기
    xml = addToList(xml, 'paraProperties', paraPrXml(ids.pp.center, 'CENTER') + paraPrXml(ids.pp.right, 'RIGHT') + paraPrXml(ids.pp.left, 'LEFT') + paraPrXml(ids.pp.title, 'LEFT', 400, true), 4);
    return { xml, ids };
  }

  // ------------------------------------------------------------------
  // section0.xml: 본문
  // ------------------------------------------------------------------
  let pid = 1000;
  const nextId = () => String(++pid);

  // newPage: 이 문단 앞에서 쪽을 나눔 (표마다 새 쪽에서 시작)
  function para(text, ppId, cpId, newPage) {
    const t = text ? `<hp:t>${esc(text)}</hp:t>` : '<hp:t/>';
    return `<hp:p id="${nextId()}" paraPrIDRef="${ppId}" styleIDRef="0" pageBreak="${newPage ? 1 : 0}" columnBreak="0" merged="0"><hp:run charPrIDRef="${cpId}">${t}</hp:run></hp:p>`;
  }

  // 글자 수로 열 너비 비율을 정함 (숫자 열은 좁게, 구분·보기 이름 열은 넓게)
  function columnWidths(grid, headRows) {
    const ncol = grid[0].length;
    const weight = new Array(ncol).fill(4);
    grid.forEach((line, r) => {
      line.forEach((c, j) => {
        if (c.merged || (c.cs && c.cs > 1)) return;
        const text = c.t === 'n' ? '100.0' : String(c.v == null ? '' : c.v);
        const len = Array.from(text).reduce((a, ch) => a + (/[ㄱ-힝]/.test(ch) ? 1 : 0.55), 0);
        // 머리글은 두세 줄로 줄바꿈되므로 절반만 반영
        const w = r < headRows && c.t === 'h' ? Math.min(len / 2 + 1, 8) : Math.min(len + 1.2, 16);
        weight[j] = Math.max(weight[j], w);
      });
    });
    const sum = weight.reduce((a, b) => a + b, 0);
    const widths = weight.map((w) => Math.floor((w / sum) * TEXT_WIDTH));
    widths[0] += TEXT_WIDTH - widths.reduce((a, b) => a + b, 0);
    return widths;
  }

  // 글자 폭 대략값(HWPUNIT): 한글·전각 = 글자 크기, 영문·숫자 = 0.55배
  function textWidth(text, size) {
    return Array.from(String(text == null ? '' : text)).reduce((a, ch) => a + (/[ㄱ-힝　-ヿ一-鿿■⊙※]/.test(ch) ? 1 : 0.55), 0) * size * 100;
  }

  // 칸 안 글자가 몇 줄이 되는지 (단어 단위로 줄바꿈되므로 약간 여유를 둠)
  function lineCount(text, size, width) {
    const avail = Math.max(width - CELL_MARGIN.left - CELL_MARGIN.right, size * 100);
    const tw = textWidth(text, size);
    return Math.max(1, Math.ceil((tw * 1.08) / avail));
  }

  const rowHeightOf = (size, lines) => Math.round(lines * size * 100 * LINE + CELL_MARGIN.top + CELL_MARGIN.bottom + 60);

  // 표 모양을 글자 크기별로 계산해 한 쪽에 들어가는 가장 큰 크기를 고름
  function layoutTable(table, opts, ids) {
    const grid = TG.toGrid(table, opts);
    const headRows = (opts.orientation || 'row') === 'row' ? 1 : 2;
    const widths = columnWidths(grid, headRows);
    const cellText = (c) => (c.t === 'n' ? TG.fmtNumber(c.v, c.fmt, opts) : c.v);
    const spanWidth = (j, cs) => widths.slice(j, j + cs).reduce((a, b) => a + b, 0);
    // 제목 한 줄 + 제목 아래 간격 + 주석 줄 수
    const extra = () => {
      const titleLines = lineCount(`표 ${table.no}. ${table.title}`, TITLE_SIZE, TEXT_WIDTH);
      const noteLines = table.notes.reduce((a, n) => a + lineCount('※ ' + n, NOTE_SIZE, TEXT_WIDTH), 0);
      return titleLines * TITLE_SIZE * 100 * LINE + 400 + noteLines * NOTE_SIZE * 100 * LINE + 600;
    };
    let chosen = null;
    for (const z of ids.sizes) {
      // 줄마다 가장 많이 줄바꿈되는 칸 기준 높이 (여러 줄에 걸친 칸은 나눠서 계산)
      const heights = grid.map((line) => {
        let lines = 1;
        line.forEach((c, j) => {
          if (c.merged) return;
          const rs = c.rs || 1;
          lines = Math.max(lines, Math.ceil(lineCount(cellText(c), z.size, spanWidth(j, c.cs || 1)) / rs));
        });
        return rowHeightOf(z.size, lines);
      });
      const total = heights.reduce((a, b) => a + b, 0) + extra();
      chosen = { z, heights, total, fits: total <= TEXT_HEIGHT };
      if (chosen.fits) break;
    }
    return { grid, headRows, widths, cellText, ...chosen };
  }

  function tableXml(table, opts, ids) {
    const L = layoutTable(table, opts, ids);
    const { grid, headRows, widths, z, heights } = L;
    const nrow = grid.length;
    const ncol = widths.length;
    let rowsXml = '';
    grid.forEach((line, r) => {
      let tr = '<hp:tr>';
      line.forEach((c, j) => {
        if (c.merged) return; // 합쳐진 칸에 덮인 칸은 적지 않음
        const cs = c.cs || 1;
        const rs = c.rs || 1;
        const w = widths.slice(j, j + cs).reduce((a, b) => a + b, 0);
        const h = heights.slice(r, r + rs).reduce((a, b) => a + b, 0);
        let bf = ids.bf.plain;
        if (r < headRows || c.t === 'h') bf = ids.bf.head;
        else if (c.total) bf = ids.bf.total;
        else if (c.t === 'g') bf = ids.bf.group;
        else if (c.stat) bf = ids.bf.stat;
        const bold = r < headRows || c.t === 'h' || c.t === 'g' || c.total;
        const cp = bold ? z.bold : z.cell;
        const pp = c.t === 'n' ? ids.pp.right : c.t === 'h' || r < headRows ? ids.pp.center : ids.pp.left;
        tr +=
          `<hp:tc name="" header="${r < headRows ? 1 : 0}" hasMargin="1" protect="0" editable="0" dirty="0" borderFillIDRef="${bf}">` +
          '<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">' +
          para(L.cellText(c), pp, cp) +
          '</hp:subList>' +
          `<hp:cellAddr colAddr="${j}" rowAddr="${r}"/><hp:cellSpan colSpan="${cs}" rowSpan="${rs}"/>` +
          `<hp:cellSz width="${w}" height="${h}"/>` +
          `<hp:cellMargin left="${CELL_MARGIN.left}" right="${CELL_MARGIN.right}" top="${CELL_MARGIN.top}" bottom="${CELL_MARGIN.bottom}"/></hp:tc>`;
      });
      rowsXml += tr + '</hp:tr>';
    });
    const height = heights.reduce((a, b) => a + b, 0);
    // 글자처럼 취급하지 않음(treatAsChar=0): 위아래로 본문과 배치, 쪽 경계에서 셀 단위로 나누고 제목 줄 반복
    const tbl =
      `<hp:tbl id="${nextId()}" zOrder="0" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="CELL" repeatHeader="1" rowCnt="${nrow}" colCnt="${ncol}" cellSpacing="0" borderFillIDRef="${ids.bf.plain}" noAdjust="0">` +
      `<hp:sz width="${TEXT_WIDTH}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/>` +
      '<hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>' +
      '<hp:outMargin left="0" right="0" top="0" bottom="283"/>' +
      `<hp:inMargin left="${CELL_MARGIN.left}" right="${CELL_MARGIN.right}" top="${CELL_MARGIN.top}" bottom="${CELL_MARGIN.bottom}"/>` +
      rowsXml +
      '</hp:tbl>';
    return `<hp:p id="${nextId()}" paraPrIDRef="${ids.pp.left}" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="${z.cell}">${tbl}<hp:t/></hp:run></hp:p>`;
  }

  function sectionBody(result, opts, ids) {
    let body = '';
    body += para('설문 통계표', ids.pp.left, ids.cp.cover);
    const info = [
      `생성일: ${new Date().toLocaleString('ko-KR')}`,
      `전체 사례수: ${result.n.toLocaleString('ko-KR')}명`,
      `배너: ${opts.banners && opts.banners.length ? opts.banners.map((b) => b.label).join(', ') : '없음(전체만)'}`,
      `가중치: ${opts.weightVar ? opts.weightVar + ' 적용' : '적용 안 함'}`,
    ];
    info.forEach((t) => (body += para(t, ids.pp.left, ids.cp.note)));
    result.tables.forEach((t) => {
      // 한 쪽에 표 하나: 표 제목마다 새 쪽에서 시작 (첫 쪽은 생성 정보)
      body += para(`표 ${t.no}. ${t.title}`, ids.pp.title, ids.cp.title, true);
      if (t.kind === 'error') {
        body += para(t.notes.join(' '), ids.pp.left, ids.cp.note);
        return;
      }
      body += tableXml(t, opts, ids);
      t.notes.forEach((n) => (body += para('※ ' + n, ids.pp.left, ids.cp.note)));
    });
    return body;
  }

  // ------------------------------------------------------------------
  function buildHwpxFiles(result, opts) {
    const T = TG.HWPX_TEMPLATE;
    if (!T) throw new Error('한글 문서 틀(hwpx-template.js)이 없습니다.');
    pid = 1000;
    const { xml: header, ids } = buildHeader(T['Contents/header.xml']);
    const sec = T['Contents/section0.xml'].replace(
      /<hp:margin header="\d+" footer="\d+" gutter="0" left="\d+" right="\d+" top="\d+" bottom="\d+"\/>/,
      `<hp:margin header="${PAGE.header}" footer="${PAGE.footer}" gutter="0" left="${PAGE.left}" right="${PAGE.right}" top="${PAGE.top}" bottom="${PAGE.bottom}"/>`,
    );
    if (!sec.includes(`left="${PAGE.left}" right="${PAGE.right}"`)) throw new Error('한글 틀의 쪽 여백을 찾지 못했습니다.');
    const end = sec.lastIndexOf('</hs:sec>');
    const section = sec.slice(0, end) + sectionBody(result, opts, ids) + sec.slice(end);
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const hpf = T['Contents/content.hpf']
      .replace(/<opf:title\/>/, '<opf:title>설문 통계표</opf:title>')
      .replace(/synthetic-fixture-author/g, '설문 통계표 생성기')
      .replace(/(name="CreatedDate" content="text">)[^<]*/, `$1${now}`)
      .replace(/(name="ModifiedDate" content="text">)[^<]*/, `$1${now}`)
      .replace(/(name="date" content="text">)[^<]*/, `$1${new Date().toLocaleString('ko-KR')}`);
    const preview = result.tables.slice(0, 20).map((t) => `표 ${t.no}. ${t.title}`).join('\r\n');
    return {
      mimetype: T.mimetype,
      'version.xml': T['version.xml'],
      'settings.xml': T['settings.xml'],
      'META-INF/container.xml': T['META-INF/container.xml'],
      'META-INF/container.rdf': T['META-INF/container.rdf'],
      'META-INF/manifest.xml': T['META-INF/manifest.xml'],
      'Contents/content.hpf': hpf,
      'Contents/header.xml': header,
      'Contents/section0.xml': section,
      'Preview/PrvText.txt': preview,
    };
  }

  async function buildHwpx(result, opts) {
    const JSZip = root.JSZip;
    const files = buildHwpxFiles(result, opts);
    const zip = new JSZip();
    // mimetype은 압축하지 않고 맨 앞에 넣어야 한글이 HWPX로 알아봄
    zip.file('mimetype', files.mimetype, { compression: 'STORE' });
    Object.keys(files).forEach((name) => {
      if (name !== 'mimetype') zip.file(name, files[name]);
    });
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/hwp+zip' });
  }

  async function downloadHwpx(result, opts, filename) {
    const bytes = await buildHwpx(result, opts);
    const blob = new Blob([bytes], { type: 'application/hwp+zip' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename || '통계표.hwpx';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  Object.assign(TG, { buildHwpxFiles, buildHwpx, downloadHwpx, hwpxLayout: layoutTable, HWPX_PAGE: { ...PAGE, TEXT_WIDTH, TEXT_HEIGHT } });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

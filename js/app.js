/*
 * 화면 동작: 파일 읽기 → 문항 확인 → 표 설정 → 결과/엑셀 저장
 */
(function () {
  'use strict';
  const TG = window.TG;
  const $ = (id) => document.getElementById(id);
  const esc = TG.esc;

  const state = {
    files: { data: null, cb: null }, // { name, wb }
    data: null,
    codebook: null,
    items: [],
    bannerOrder: [], // 선택 순서대로 item id
    bannerNames: {},
    result: null,
    opts: null,
    derived: [], // 가공 변수 설정 (파일을 다시 읽어도 유지)
    editing: null, // 고치는 중인 가공 변수 id (새로 만들 때는 'new')
  };

  // ------------------------------------------------------------------
  // 파일 읽기
  // ------------------------------------------------------------------
  function decodeCsv(buf) {
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, '');
    } catch (e) {
      return new TextDecoder('euc-kr').decode(buf); // 한글 윈도우 CSV
    }
  }

  function readWorkbook(name, buf) {
    if (/\.csv$/i.test(name)) return XLSX.read(decodeCsv(buf), { type: 'string', raw: true });
    return XLSX.read(buf, { type: 'array' });
  }

  function sheetRows(wb, sheet) {
    return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null, raw: true, blankrows: false });
  }

  function setFile(kind, name, buf) {
    const wb = readWorkbook(name, buf);
    state.files[kind] = { name, wb };
    const isData = kind === 'data';
    $(isData ? 'nameData' : 'nameCb').textContent = name;
    $(isData ? 'dropData' : 'dropCb').classList.add('ok');
    const sel = $(isData ? 'sheetData' : 'sheetCb');
    sel.innerHTML = wb.SheetNames.map((s) => `<option>${esc(s)}</option>`).join('');
    if (!isData) {
      const guess = wb.SheetNames.find((s) => /코드|code/i.test(s));
      if (guess) sel.value = guess;
    }
    sel.classList.toggle('hidden', wb.SheetNames.length < 2);
    if (!isData) updateFormatInfo();
    $('btnAnalyze').disabled = !state.files.data;
    hideMsg();
  }

  function onFile(kind, file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        setFile(kind, file.name, new Uint8Array(reader.result));
      } catch (e) {
        showMsg(`${file.name} 파일을 읽지 못했습니다: ${e.message}`, true);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function bindDrop(dropId, inputId, kind) {
    const drop = $(dropId);
    $(inputId).addEventListener('change', (e) => onFile(kind, e.target.files[0]));
    ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
    drop.addEventListener('drop', (e) => onFile(kind, e.dataTransfer.files[0]));
    // 상자를 누르면 파일 선택 창 열기 (시트 선택 상자는 제외)
    drop.addEventListener('click', (e) => {
      if (['SELECT', 'OPTION', 'INPUT'].includes(e.target.tagName)) return;
      $(inputId).click();
    });
  }

  // ------------------------------------------------------------------
  // 코드북 형식
  // ------------------------------------------------------------------
  function codebookSheets() {
    const wb = state.files.cb.wb;
    const sheets = {};
    wb.SheetNames.forEach((n) => (sheets[n] = sheetRows(wb, n)));
    return sheets;
  }

  function chosenFormat() {
    if (!state.files.cb) return null;
    const v = $('cbFormat').value;
    if (v !== 'auto') return TG.FORMATS.find((f) => f.id === v);
    return TG.detectFormat(codebookSheets());
  }

  function updateFormatInfo() {
    const fmt = chosenFormat();
    if (!fmt) {
      $('fmtInfo').textContent = '';
      return;
    }
    const auto = $('cbFormat').value === 'auto';
    $('fmtInfo').textContent = `${auto ? '감지된 형식' : '선택한 형식'}: ${fmt.name}${fmt.usesSheet ? '' : ' (모든 시트를 함께 읽음)'}`;
    const sel = $('sheetCb');
    sel.classList.toggle('hidden', !fmt.usesSheet || state.files.cb.wb.SheetNames.length < 2);
  }

  function showMsg(text, isError) {
    const m = $('msg1');
    m.textContent = text;
    m.classList.toggle('error', !!isError);
    m.classList.remove('hidden');
  }
  function hideMsg() {
    $('msg1').classList.add('hidden');
  }

  async function loadSamples() {
    try {
      const get = async (f) => {
        const res = await fetch('samples/' + encodeURIComponent(f));
        if (!res.ok) throw new Error(res.status);
        return new Uint8Array(await res.arrayBuffer());
      };
      setFile('data', '예시_데이터.xlsx', await get('예시_데이터.xlsx'));
      setFile('cb', '예시_코드북.xlsx', await get('예시_코드북.xlsx'));
      analyze();
    } catch (e) {
      showMsg('예시 파일을 불러오지 못했습니다. (내 PC에서 index.html을 직접 연 경우) 오른쪽 링크로 예시 파일을 내려받아 올려 주세요.', true);
    }
  }

  // ------------------------------------------------------------------
  // 2단계: 문항 자동 인식
  // ------------------------------------------------------------------
  function analyze() {
    try {
      const d = state.files.data;
      state.data = TG.prepareData(sheetRows(d.wb, $('sheetData').value || d.wb.SheetNames[0]));
      const warnings = [];
      if (state.files.cb) {
        const c = state.files.cb;
        state.codebook = TG.parseCodebookBook(codebookSheets(), { format: $('cbFormat').value, sheet: $('sheetCb').value || c.wb.SheetNames[0] });
        warnings.push(...state.codebook.warnings);
      } else {
        state.codebook = null;
        warnings.push('코드북 없이 데이터만으로 인식했습니다. 보기 이름이 코드 번호로 표시됩니다.');
      }
      const built = TG.buildItems(state.codebook, state.data);
      warnings.push(...built.warnings);
      state.items = built.items;
      const dv = TG.applyDerived(state.data, state.items, state.derived);
      state.items = dv.items;
      warnings.push(...dv.warnings);
      state.bannerOrder = [];
      state.bannerNames = {};
      closeDeriveForm();
      renderWarnings(warnings);
      renderDerived();
      renderItems();
      renderOptions();
      $('step2').classList.remove('hidden');
      $('step3').classList.remove('hidden');
      $('step4').classList.add('hidden');
      const fmtText = state.codebook ? ` 코드북 형식: ${state.codebook.formatName}.` : '';
      showMsg(`응답자 ${state.data.n.toLocaleString('ko-KR')}명, 변수 ${state.data.order.length}개를 읽었습니다.${fmtText}`);
      $('step2').scrollIntoView({ behavior: 'smooth' });
    } catch (e) {
      console.error(e);
      showMsg('파일을 해석하지 못했습니다: ' + e.message, true);
    }
  }

  function renderWarnings(list) {
    const box = $('warnings');
    box.innerHTML = list.length ? '<ul>' + list.map((w) => `<li>${esc(w)}</li>`).join('') + '</ul>' : '';
    box.classList.toggle('hidden', !list.length);
  }

  function itemDetail(it) {
    if (it.isGroup) {
      const lab = it.itemLabels.join(', ');
      const codes = it.kind === 'multi01' ? '' : ` · 보기 ${it.codes.length}개`;
      return `항목 ${it.itemLabels.length}개: ${lab}${codes}`;
    }
    if (!it.codes.length) return '';
    return it.codes.map((c) => `${c.code}=${c.label}`).join(', ');
  }

  function renderItems() {
    const tbody = $('itemTable').querySelector('tbody');
    tbody.innerHTML = state.items
      .map((it) => {
        const kinds = it.isGroup ? TG.GROUP_KINDS : TG.SINGLE_KINDS;
        const opts = kinds.map((k) => `<option value="${k}"${k === it.kind ? ' selected' : ''}>${TG.KINDS[k]}</option>`).join('');
        const vars = it.vars.length > 3 ? `${it.vars[0]} ~ ${it.vars[it.vars.length - 1]} (${it.vars.length}개)` : it.vars.join(', ');
        return `<tr data-id="${it.id}" class="${it.include ? '' : 'off'}">
          <td><input type="checkbox" class="inc"${it.include ? ' checked' : ''}></td>
          <td class="vars">${esc(vars)}</td>
          <td><input type="text" class="title" value="${esc(it.title)}"></td>
          <td><select class="kind${it.kind !== it.autoKind ? ' changed' : ''}">${opts}</select></td>
          <td class="sort">${TG.canSort(it) ? `<input type="checkbox" class="srt"${it.sort ? ' checked' : ''} aria-label="큰 순 정렬">` : '<span class="na" title="수치·척도 문항은 보기 순서를 유지합니다">–</span>'}</td>
          <td class="reason">${esc(it.reason)}</td>
          <td class="labels" title="${esc(itemDetail(it))}">${esc(truncate(itemDetail(it), 70))}</td>
        </tr>`;
      })
      .join('');
    updateSummary();
  }

  function truncate(s, n) {
    return s.length > n ? s.slice(0, n) + '…' : s;
  }

  function updateSummary() {
    const on = state.items.filter((i) => i.include).length;
    $('itemSummary').textContent = `문항 ${state.items.length}개 중 ${on}개 선택`;
  }

  function itemById(id) {
    return state.items.find((i) => i.id === id);
  }

  function bindItemTable() {
    const tbody = $('itemTable').querySelector('tbody');
    tbody.addEventListener('change', (e) => {
      const tr = e.target.closest('tr');
      if (!tr) return;
      const it = itemById(tr.dataset.id);
      if (e.target.classList.contains('inc')) {
        it.include = e.target.checked;
        tr.classList.toggle('off', !it.include);
      } else if (e.target.classList.contains('srt')) {
        it.sort = e.target.checked;
      } else if (e.target.classList.contains('kind')) {
        it.kind = e.target.value;
        e.target.classList.toggle('changed', it.kind !== it.autoKind);
        if (['exclude', 'weight'].includes(it.kind)) {
          it.include = false;
        } else {
          it.include = true;
        }
        tr.querySelector('.inc').checked = it.include;
        tr.classList.toggle('off', !it.include);
        tr.querySelector('td.sort').innerHTML = TG.canSort(it)
          ? `<input type="checkbox" class="srt"${it.sort ? ' checked' : ''} aria-label="큰 순 정렬">`
          : '<span class="na" title="수치·척도 문항은 보기 순서를 유지합니다">–</span>';
        renderOptions();
      }
      updateSummary();
    });
    tbody.addEventListener('input', (e) => {
      if (!e.target.classList.contains('title')) return;
      itemById(e.target.closest('tr').dataset.id).title = e.target.value;
    });
    $('checkAll').addEventListener('click', () => setAll(true));
    $('uncheckAll').addEventListener('click', () => setAll(false));
    $('sortAll').addEventListener('click', () => setSortAll(true));
    $('unsortAll').addEventListener('click', () => setSortAll(false));
  }

  function setSortAll(on) {
    state.items.forEach((it) => {
      if (TG.canSort(it)) it.sort = on;
    });
    renderItems();
  }

  function setAll(on) {
    state.items.forEach((it) => {
      if (on && ['exclude', 'weight'].includes(it.kind)) return;
      it.include = on;
    });
    renderItems();
  }

  // ------------------------------------------------------------------
  // 2단계: 가공 변수 (값 계산 / 구간 나누기 / 보기 묶기)
  // ------------------------------------------------------------------
  let deriveSeq = 0;
  const newDeriveId = () => 'v' + Date.now().toString(36) + (++deriveSeq);
  const dfType = () => document.querySelector('input[name="dfType"]:checked').value;
  const itemOfKey = (key) => state.items.find((it) => !it.isGroup && it.keys[0] === key);

  function applyDerivedNow() {
    const r = TG.applyDerived(state.data, state.items, state.derived);
    state.items = r.items;
    renderDerived(r.warnings);
    renderItems();
    renderOptions();
    return r.warnings;
  }

  function renderDerived(warnings) {
    const bad = new Set((warnings || []).map((w) => (/^가공 변수 (\S+):/.exec(w) || [])[1]));
    $('deriveList').innerHTML = state.derived
      .map((d) => {
        const made = state.data && state.data.columns.get(TG.keyOf(d.name));
        const ok = made && made.derived && made.derived.id === d.id;
        let text;
        try {
          text = TG.describeDerived(d);
        } catch (e) {
          text = '';
        }
        return `<div class="d${ok ? '' : ' bad'}" data-id="${d.id}">
          <code>${esc(d.name)}</code>
          <span class="t">${esc(TG.DERIVE_TYPES[d.type])}</span>
          <span class="x">${esc(text)}${ok ? '' : ' · <b>만들지 못함</b>(원래 변수를 확인하세요)'}${bad.has(d.name) && ok ? ' · 확인 필요' : ''}</span>
          <span><button class="btn tiny" data-act="edit">고치기</button> <button class="btn tiny" data-act="del">삭제</button></span>
        </div>`;
      })
      .join('');
    $('btnSaveDerive').disabled = !state.derived.length;
  }

  // 원래 변수 후보: 값 계산·구간 나누기 → 숫자 변수, 보기 묶기 → 보기가 있는 단일 문항
  function sourceOptions(type) {
    const idx = state.derived.findIndex((d) => d.id === state.editing);
    // 고치는 중인 변수와 그 뒤에 만든 가공 변수는 원래 변수로 쓸 수 없음
    const blocked = new Set(idx >= 0 ? state.derived.slice(idx).map((d) => TG.keyOf(d.name)) : []);
    const list = [];
    state.data.order.forEach((key) => {
      if (blocked.has(key)) return;
      const col = state.data.columns.get(key);
      const it = itemOfKey(key);
      if (type === 'merge') {
        if (!it || it.codes.length < 2) return;
      } else if (col.isText || !col.filled) return;
      const title = it ? it.title : '';
      list.push({ key, name: col.name, text: title && title !== col.name ? `${col.name} · ${title.slice(0, 40)}` : col.name, title });
    });
    return list;
  }

  function fillSources(keep) {
    const type = dfType();
    const list = sourceOptions(type);
    const sel = $('dfSource');
    const prev = keep || sel.value;
    sel.innerHTML = list.map((o) => `<option value="${esc(o.name)}">${esc(o.text)}</option>`).join('');
    if (prev && list.some((o) => o.name === prev)) sel.value = prev;
  }

  function syncDeriveForm() {
    const type = dfType();
    document.querySelectorAll('#deriveForm .df-compute').forEach((el) => el.classList.toggle('hidden', type !== 'compute'));
    document.querySelectorAll('#deriveForm .df-bin').forEach((el) => el.classList.toggle('hidden', type !== 'bin'));
    document.querySelectorAll('#deriveForm .df-merge').forEach((el) => el.classList.toggle('hidden', type !== 'merge'));
  }

  function fillMerge(groups) {
    const codes = TG.sourceCodes(state.items, TG.keyOf($('dfSource').value));
    $('dfMerge').querySelector('tbody').innerHTML = codes
      .map((c) => {
        const v = groups && groups[c.code] != null ? groups[c.code] : c.label;
        return `<tr data-code="${c.code}"><td>${esc(c.code + '. ' + c.label)}</td><td><input type="text" value="${esc(v)}"></td></tr>`;
      })
      .join('');
  }

  // 이름·제목을 직접 고치지 않았으면 원래 변수에 맞춰 자동으로 채움
  function autoNames() {
    const type = dfType();
    const src = $('dfSource').value;
    if (!src) return;
    const it = itemOfKey(TG.keyOf(src));
    const title = it ? TG.shortName(it.title, src) : src;
    if (!$('dfName').dataset.touched) $('dfName').value = src.replace(/[^A-Za-z0-9가-힣_]/g, '_') + { compute: '_C', bin: '_G', merge: '_M' }[type];
    if (!$('dfLabel').dataset.touched) {
      let label = title;
      if (type === 'compute' && $('dfK').value.trim() !== '') label = `${title} (${TG.describeDerived({ type, op: $('dfOp').value, k: $('dfK').value, source: src })})`;
      else if (type === 'bin') label = `${title} (구간)`;
      else if (type === 'merge') label = `${title} (묶음)`;
      $('dfLabel').value = label;
    }
  }

  function readDeriveForm() {
    const type = dfType();
    const def = { id: state.editing === 'new' ? newDeriveId() : state.editing, type, name: $('dfName').value.trim(), label: $('dfLabel').value.trim(), source: $('dfSource').value };
    if (type === 'compute') {
      def.op = $('dfOp').value;
      def.k = $('dfK').value.trim();
    } else if (type === 'bin') {
      def.bins = $('dfBins').value.trim();
    } else {
      def.groups = {};
      $('dfMerge').querySelectorAll('tbody tr').forEach((tr) => (def.groups[tr.dataset.code] = tr.querySelector('input').value.trim()));
    }
    return def;
  }

  const fmt = (n) => (n == null ? '-' : (Math.round(n * 100) / 100).toLocaleString('ko-KR'));

  // 입력하는 동안 결과를 미리 보여 줌
  function previewDerive() {
    const def = readDeriveForm();
    const others = state.derived.filter((d) => d.id !== def.id);
    const err = TG.checkDerived(def, state.data, state.items, others);
    const box = $('dfPreview');
    if (err) {
      box.innerHTML = `<span>${esc(err)}</span>`;
      return;
    }
    const r = TG.previewDerived(def, state.data, state.items);
    const vals = r.values.filter((v) => v != null);
    if (def.type === 'compute') {
      if (!vals.length) {
        box.textContent = '계산된 값이 없습니다.';
        return;
      }
      const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
      box.innerHTML = `미리보기(가중치 없이): 응답 ${vals.length.toLocaleString('ko-KR')}명 · 평균 ${fmt(mean)} · 최솟값 ${fmt(Math.min(...vals))} · 최댓값 ${fmt(Math.max(...vals))}`;
      return;
    }
    const cnt = new Map(r.codes.map((c) => [c.code, 0]));
    vals.forEach((v) => cnt.set(v, (cnt.get(v) || 0) + 1));
    const rows = r.codes.map((c) => `<tr><th>${esc(c.label)}</th><td>${cnt.get(c.code).toLocaleString('ko-KR')}명</td><td>${vals.length ? ((cnt.get(c.code) / vals.length) * 100).toFixed(1) : '-'}%</td></tr>`).join('');
    const drop = r.dropped ? ` · <b>어느 구간에도 들지 않는 값 ${r.dropped}건</b>(빈 값 처리)` : '';
    box.innerHTML = `미리보기(가중치 없이): 응답 ${vals.length.toLocaleString('ko-KR')}명${drop}<table>${rows}</table>`;
  }

  function openDeriveForm(def) {
    state.editing = def ? def.id : 'new';
    const type = def ? def.type : 'compute';
    document.querySelector(`input[name="dfType"][value="${type}"]`).checked = true;
    syncDeriveForm();
    fillSources(def ? def.source : null);
    $('dfOp').value = def && def.op ? def.op : 'k-x';
    $('dfK').value = def && def.k != null ? def.k : '';
    $('dfBins').value = def && def.bins ? def.bins : '';
    $('dfName').value = def ? def.name : '';
    $('dfLabel').value = def ? state.items.find((it) => it.derived && it.derived.id === def.id)?.title || def.label || '' : '';
    $('dfName').dataset.touched = def ? '1' : '';
    $('dfLabel').dataset.touched = def ? '1' : '';
    fillMerge(def && def.groups);
    if (!def) autoNames();
    $('dfOk').textContent = def ? '고치기' : '추가';
    $('dfError').classList.add('hidden');
    $('deriveForm').classList.remove('hidden');
    previewDerive();
    $('deriveForm').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  function closeDeriveForm() {
    state.editing = null;
    $('deriveForm').classList.add('hidden');
  }

  function saveDeriveForm() {
    const def = readDeriveForm();
    const others = state.derived.filter((d) => d.id !== def.id);
    const err = TG.checkDerived(def, state.data, state.items, others);
    if (err) {
      $('dfError').textContent = err;
      $('dfError').classList.remove('hidden');
      return;
    }
    const idx = state.derived.findIndex((d) => d.id === def.id);
    if (idx >= 0) state.derived[idx] = def;
    else state.derived.push(def);
    // 표 제목을 바꿨으면 문항 목록의 제목도 바꿈
    const it = state.items.find((x) => x.derived && x.derived.id === def.id);
    if (it) it.derivedLabel = null;
    closeDeriveForm();
    applyDerivedNow();
  }

  function saveDeriveFile() {
    const body = JSON.stringify({ app: '설문 통계표 생성기', kind: '가공 설정', version: 1, derived: state.derived }, null, 2);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([body], { type: 'application/json' }));
    const base = state.files.data ? state.files.data.name.replace(/\.[^.]+$/, '') : '데이터';
    a.download = `${base}_가공설정.json`;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  function loadDeriveFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const obj = JSON.parse(reader.result);
        const list = Array.isArray(obj) ? obj : obj.derived;
        if (!Array.isArray(list)) throw new Error('가공 설정 파일이 아닙니다.');
        const names = new Set(list.map((d) => TG.keyOf(d.name)));
        state.derived = state.derived.filter((d) => !names.has(TG.keyOf(d.name))).concat(list.map((d) => ({ ...d, id: newDeriveId() })));
        closeDeriveForm();
        const w = applyDerivedNow();
        alert(`가공 변수 ${list.length}개를 불러왔습니다.${w.length ? '\n\n확인 필요:\n' + w.join('\n') : ''}`);
      } catch (e) {
        alert('가공 설정을 읽지 못했습니다: ' + e.message);
      }
      $('fileDerive').value = '';
    };
    reader.readAsText(file);
  }

  function bindDerive() {
    $('dfOp').innerHTML = Object.entries(TG.DERIVE_OPS).map(([k, o]) => `<option value="${k}">${esc(o.label)}</option>`).join('');
    $('btnAddDerive').addEventListener('click', () => openDeriveForm(null));
    $('dfCancel').addEventListener('click', closeDeriveForm);
    $('dfOk').addEventListener('click', saveDeriveForm);
    document.querySelectorAll('input[name="dfType"]').forEach((r) =>
      r.addEventListener('change', () => {
        syncDeriveForm();
        fillSources();
        fillMerge();
        autoNames();
        previewDerive();
      }),
    );
    $('dfSource').addEventListener('change', () => {
      fillMerge();
      autoNames();
      previewDerive();
    });
    ['dfName', 'dfLabel'].forEach((id) => $(id).addEventListener('input', () => ($(id).dataset.touched = '1')));
    ['dfOp', 'dfK', 'dfBins', 'dfName'].forEach((id) =>
      $(id).addEventListener('input', () => {
        if (id === 'dfOp' || id === 'dfK') autoNames();
        previewDerive();
      }),
    );
    $('dfOp').addEventListener('change', () => {
      autoNames();
      previewDerive();
    });
    $('dfMerge').addEventListener('input', previewDerive);
    $('deriveList').addEventListener('click', (e) => {
      const btn = e.target.closest('button');
      if (!btn) return;
      const id = btn.closest('.d').dataset.id;
      const def = state.derived.find((d) => d.id === id);
      if (btn.dataset.act === 'edit') openDeriveForm(def);
      else if (btn.dataset.act === 'del') {
        if (!confirm(`가공 변수 ${def.name}을(를) 삭제할까요?`)) return;
        state.derived = state.derived.filter((d) => d.id !== id);
        if (state.editing === id) closeDeriveForm();
        applyDerivedNow();
      }
    });
    $('btnSaveDerive').addEventListener('click', saveDeriveFile);
    $('btnLoadDerive').addEventListener('click', () => $('fileDerive').click());
    $('fileDerive').addEventListener('change', (e) => loadDeriveFile(e.target.files[0]));
  }

  // ------------------------------------------------------------------
  // 3단계: 배너 / 가중치 / 표 모양
  // ------------------------------------------------------------------
  function bannerCandidates() {
    return state.items.filter((it) => !it.isGroup && ['single', 'scale'].includes(it.kind) && it.codes.length >= 2 && it.codes.length <= 30);
  }

  function renderOptions() {
    const list = bannerCandidates();
    state.bannerOrder = state.bannerOrder.filter((id) => list.some((it) => it.id === id));
    $('bannerList').innerHTML = list.length
      ? list
          .map((it) => {
            const idx = state.bannerOrder.indexOf(it.id);
            const name = state.bannerNames[it.id] || TG.shortName(it.title, it.vars[0]);
            return `<div class="b${idx >= 0 ? ' on' : ''}" data-id="${it.id}">
            <input type="checkbox"${idx >= 0 ? ' checked' : ''} aria-label="배너로 사용">
            <code title="${esc(it.vars[0])}">${idx >= 0 ? `<span class="order">${idx + 1}</span>` : ''}${esc(it.vars[0])}</code>
            <input type="text" value="${esc(name)}" aria-label="배너 이름">
            <span class="cnt">${it.codes.length}개</span>
          </div>`;
          })
          .join('')
      : '<div class="b">배너로 쓸 수 있는 단일응답 문항이 없습니다.</div>';

    // 가중치 후보: 숫자 변수
    const sel = $('weightVar');
    const prev = sel.value;
    const numeric = state.data.order.map((k) => state.data.columns.get(k)).filter((c) => !c.isText && c.filled > 0);
    const weightItem = state.items.find((i) => i.kind === 'weight');
    sel.innerHTML = numeric.map((c) => `<option value="${esc(c.name)}">${esc(c.name)}</option>`).join('');
    if (prev && numeric.some((c) => c.name === prev)) sel.value = prev;
    else if (weightItem) {
      sel.value = weightItem.vars[0];
      $('useWeight').checked = true;
    } else $('useWeight').checked = false;
    syncWeight();
  }

  function syncWeight() {
    const on = $('useWeight').checked;
    $('weightVar').disabled = !on;
    $('baseModeBox').classList.toggle('hidden', !on);
  }

  function bindOptions() {
    const box = $('bannerList');
    box.addEventListener('change', (e) => {
      const row = e.target.closest('.b');
      if (!row || e.target.type !== 'checkbox') return;
      const id = row.dataset.id;
      if (e.target.checked) state.bannerOrder.push(id);
      else state.bannerOrder = state.bannerOrder.filter((x) => x !== id);
      renderOptions();
    });
    box.addEventListener('input', (e) => {
      const row = e.target.closest('.b');
      if (row && e.target.type === 'text') state.bannerNames[row.dataset.id] = e.target.value;
    });
    $('useWeight').addEventListener('change', syncWeight);
  }

  function readOptions() {
    const radio = (name) => document.querySelector(`input[name="${name}"]:checked`).value;
    const useW = $('useWeight').checked && $('weightVar').value;
    return {
      banners: state.bannerOrder.map((id) => {
        const it = itemById(id);
        return { var: it.vars[0], label: state.bannerNames[id] || TG.shortName(it.title, it.vars[0]), codes: it.codes };
      }),
      weightVar: useW ? $('weightVar').value : null,
      baseMode: radio('baseMode'),
      orientation: radio('orient'),
      decimals: Number($('decimals').value),
      showSum: $('showSum').checked,
      sheetMode: $('sheetMode').value,
    };
  }

  // ------------------------------------------------------------------
  // 4단계: 결과
  // ------------------------------------------------------------------
  function run() {
    const opts = readOptions();
    if (opts.weightVar) {
      const w = state.data.columns.get(TG.keyOf(opts.weightVar)).values;
      const bad = w.filter((v) => v == null || v <= 0).length;
      if (bad && !confirm(`가중치 변수(${opts.weightVar})에 비어 있거나 0 이하인 값이 ${bad}건 있습니다. 해당 응답자는 가중 계산에서 제외됩니다. 계속할까요?`)) return;
    }
    const btn = $('btnRun');
    btn.disabled = true;
    btn.textContent = '계산 중…';
    setTimeout(() => {
      try {
        state.opts = opts;
        state.result = TG.computeTables(state.items, state.data, opts);
        renderResult();
      } catch (e) {
        console.error(e);
        alert('표를 만드는 중 오류가 났습니다: ' + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = '표 만들기';
      }
    }, 20);
  }

  function renderResult() {
    const { tables } = state.result;
    const opts = state.opts;
    $('resultInfo').textContent = `표 ${tables.length}개 · 응답자 ${state.result.n.toLocaleString('ko-KR')}명${opts.weightVar ? ' · 가중치 ' + opts.weightVar : ''}`;
    $('toc').innerHTML = tables
      .map((t) => `<a href="#t${t.no}" data-no="${t.no}" data-find="${esc((t.no + ' ' + t.title + ' ' + t.vars.join(' ')).toLowerCase())}"><b>${t.no}</b>${esc(t.title)}</a>`)
      .join('');
    $('tables').innerHTML = tables.length ? tables.map((t) => TG.tableHTML(t, opts)).join('') : '<p>선택된 문항이 없습니다.</p>';
    $('tocSearch').value = '';
    filterToc();
    $('step4').classList.remove('hidden');
    updateBarHeight();
    watchTables();
    $('step4').scrollIntoView({ behavior: 'smooth' });
  }

  // ------------------------------------------------------------------
  // 결과 화면 이동 편의 기능
  // ------------------------------------------------------------------
  // 위에 붙는 도구 모음 높이만큼 목차·표 위치를 내림
  function updateBarHeight() {
    const h = $('resultBar').offsetHeight || 64;
    document.documentElement.style.setProperty('--bar-h', h + 'px');
  }

  // 표 찾기: 번호("12"), 제목 일부, 변수명으로 목록을 줄임
  function filterToc() {
    const q = $('tocSearch').value.trim().toLowerCase();
    const links = $('toc').querySelectorAll('a');
    let shown = 0;
    links.forEach((a) => {
      const hit = !q || (/^\d+$/.test(q) ? a.dataset.no === q || a.dataset.find.includes(q) : a.dataset.find.includes(q));
      a.hidden = !hit;
      if (hit) shown++;
    });
    let empty = $('toc').querySelector('.empty');
    if (!shown && links.length) {
      if (!empty) {
        empty = document.createElement('div');
        empty.className = 'empty';
        empty.textContent = '맞는 표가 없습니다.';
        $('toc').appendChild(empty);
      }
    } else if (empty) empty.remove();
    $('tocCount').textContent = q ? `${links.length}개 중 ${shown}개` : `표 ${links.length}개`;
  }

  function goToTable(no) {
    const el = document.getElementById('t' + no);
    if (!el) return;
    el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    el.classList.add('flash');
    setTimeout(() => el.classList.remove('flash'), 1200);
    history.replaceState(null, '', '#t' + no);
  }

  // 지금 보고 있는 표를 목차에 표시하고, 도구 모음에 제목을 보여 줌
  let observer = null;
  function watchTables() {
    if (observer) observer.disconnect();
    if (!('IntersectionObserver' in window)) return;
    const visible = new Map();
    observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((e) => (e.isIntersecting ? visible.set(e.target.id, e.target) : visible.delete(e.target.id)));
        const top = Array.from(visible.values()).sort((a, b) => a.getBoundingClientRect().top - b.getBoundingClientRect().top)[0];
        if (top) setActive(top.id.slice(1));
      },
      { rootMargin: `-${($('resultBar').offsetHeight || 64) + 4}px 0px -55% 0px` },
    );
    document.querySelectorAll('#tables .tbl').forEach((el) => observer.observe(el));
  }

  function setActive(no) {
    const toc = $('toc');
    const prev = toc.querySelector('a.active');
    if (prev && prev.dataset.no === no) return;
    if (prev) prev.classList.remove('active');
    const a = toc.querySelector(`a[data-no="${no}"]`);
    if (!a) return;
    a.classList.add('active');
    // 목차 안에서만 스크롤해서 현재 표가 보이게 함 (페이지는 움직이지 않음)
    const list = toc;
    const top = a.offsetTop - list.offsetTop;
    if (top < list.scrollTop || top > list.scrollTop + list.clientHeight - a.offsetHeight) {
      list.scrollTop = top - list.clientHeight / 3;
    }
    const t = state.result && state.result.tables.find((x) => String(x.no) === no);
    $('nowTable').textContent = t ? `지금: 표 ${t.no}. ${t.title}` : '';
  }

  function bindNavigation() {
    $('tocSearch').addEventListener('input', filterToc);
    $('tocSearch').addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      const first = Array.from($('toc').querySelectorAll('a')).find((a) => !a.hidden);
      if (first) goToTable(first.dataset.no);
    });
    $('toc').addEventListener('click', (e) => {
      const a = e.target.closest('a');
      if (!a) return;
      e.preventDefault();
      goToTable(a.dataset.no);
    });
    const smoothTo = (el) => el.scrollIntoView({ behavior: 'smooth', block: 'start' });
    $('goTop').addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
    $('goSettings').addEventListener('click', () => smoothTo($('step3')));
    $('goResults').addEventListener('click', () => smoothTo($('step4')));
    $('btnBackToSettings').addEventListener('click', () => smoothTo($('step3')));
    const onScroll = () => {
      const far = window.scrollY > 500;
      $('floatNav').classList.toggle('hidden', !far);
      $('goSettings').hidden = $('step3').classList.contains('hidden');
      $('goResults').hidden = $('step4').classList.contains('hidden');
    };
    window.addEventListener('scroll', onScroll, { passive: true });
    window.addEventListener('resize', () => {
      if (!$('step4').classList.contains('hidden')) {
        updateBarHeight();
        watchTables();
      }
    });
    onScroll();
  }

  async function saveAs(btnId, label, what, run) {
    if (!state.result) return;
    const btn = $(btnId);
    btn.disabled = true;
    btn.textContent = '저장 중…';
    try {
      const base = (state.files.data.name || '데이터').replace(/\.[^.]+$/, '');
      await run(`${base}_통계표`);
    } catch (e) {
      console.error(e);
      alert(`${what} 저장 중 오류가 났습니다: ` + e.message);
    } finally {
      btn.disabled = false;
      btn.textContent = label;
    }
  }

  const saveExcel = () => saveAs('btnExcel', '엑셀로 저장', '엑셀', (base) => TG.downloadExcel(state.result, state.opts, `${base}.xlsx`));
  const saveHwpx = () => saveAs('btnHwpx', '한글로 저장', '한글', (base) => TG.downloadHwpx(state.result, state.opts, `${base}.hwpx`));
  const saveSps = () =>
    saveAs('btnSps', 'SPSS 신택스', 'SPSS 신택스', async (base) => {
      const opts = { ...state.opts, sortedAny: state.items.some((it) => it.include && it.sort) };
      TG.downloadSyntax(state.result, state.items, state.data, opts, `${base}.sps`, $('spsEncoding').value);
    });

  // ------------------------------------------------------------------
  $('cbFormat').innerHTML += TG.FORMATS.map((f) => `<option value="${f.id}">${esc(f.name)}</option>`).join('');
  $('cbFormat').addEventListener('change', updateFormatInfo);
  bindNavigation();
  bindDrop('dropData', 'fileData', 'data');
  bindDrop('dropCb', 'fileCb', 'cb');
  bindItemTable();
  bindDerive();
  bindOptions();
  $('btnAnalyze').addEventListener('click', analyze);
  $('btnSample').addEventListener('click', loadSamples);
  $('btnRun').addEventListener('click', run);
  $('btnExcel').addEventListener('click', saveExcel);
  $('btnHwpx').addEventListener('click', saveHwpx);
  $('btnSps').addEventListener('click', saveSps);
})();

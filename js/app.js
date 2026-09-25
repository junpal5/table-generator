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
        state.codebook = TG.parseCodebook(sheetRows(c.wb, $('sheetCb').value || c.wb.SheetNames[0]));
        warnings.push(...state.codebook.warnings);
      } else {
        state.codebook = null;
        warnings.push('코드북 없이 데이터만으로 인식했습니다. 보기 이름이 코드 번호로 표시됩니다.');
      }
      const built = TG.buildItems(state.codebook, state.data);
      warnings.push(...built.warnings);
      state.items = built.items;
      state.bannerOrder = [];
      state.bannerNames = {};
      renderWarnings(warnings);
      renderItems();
      renderOptions();
      $('step2').classList.remove('hidden');
      $('step3').classList.remove('hidden');
      $('step4').classList.add('hidden');
      showMsg(`응답자 ${state.data.n.toLocaleString('ko-KR')}명, 변수 ${state.data.order.length}개를 읽었습니다.`);
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
  }

  function setAll(on) {
    state.items.forEach((it) => {
      if (on && ['exclude', 'weight'].includes(it.kind)) return;
      it.include = on;
    });
    renderItems();
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
            <code>${idx >= 0 ? `<span class="order">${idx + 1}</span>` : ''}${esc(it.vars[0])}</code>
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
        return { var: it.vars[0], label: state.bannerNames[id] || TG.shortName(it.title, it.vars[0]) };
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
    $('toc').innerHTML = tables.map((t) => `<a href="#t${t.no}"><b>${t.no}</b>${esc(t.title)}</a>`).join('');
    $('tables').innerHTML = tables.length ? tables.map((t) => TG.tableHTML(t, opts)).join('') : '<p>선택된 문항이 없습니다.</p>';
    $('step4').classList.remove('hidden');
    $('step4').scrollIntoView({ behavior: 'smooth' });
  }

  async function saveExcel() {
    if (!state.result) return;
    const btn = $('btnExcel');
    btn.disabled = true;
    btn.textContent = '저장 중…';
    try {
      const base = (state.files.data.name || '데이터').replace(/\.[^.]+$/, '');
      await TG.downloadExcel(state.result, state.opts, `${base}_통계표.xlsx`);
    } catch (e) {
      console.error(e);
      alert('엑셀 저장 중 오류가 났습니다: ' + e.message);
    } finally {
      btn.disabled = false;
      btn.textContent = '엑셀로 저장';
    }
  }

  // ------------------------------------------------------------------
  bindDrop('dropData', 'fileData', 'data');
  bindDrop('dropCb', 'fileCb', 'cb');
  bindItemTable();
  bindOptions();
  $('btnAnalyze').addEventListener('click', analyze);
  $('btnSample').addEventListener('click', loadSamples);
  $('btnRun').addEventListener('click', run);
  $('btnExcel').addEventListener('click', saveExcel);
})();

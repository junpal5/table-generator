/*
 * 가공 변수: 원래 데이터에서 새 변수를 만들어 표·배너에 씁니다.
 *   값 계산   (compute) : 기준값 - 값, 값 - 기준값, 값 × 기준값, 값 ÷ 기준값   예) 2026 - 설립연도 → 업력
 *   구간 나누기 (bin)    : 숫자를 구간으로 나눔 ("~4 / 5~9 / 10~")            → 분포 + 원래 값 평균
 *   보기 묶기  (merge)   : 여러 보기를 하나로 묶음 (시도 → 권역)
 * 9999·999999 같은 모름 코드는 값 계산·구간 나누기에서 빈 값으로 둡니다.
 * 같은 내용을 SPSS 신택스에도 COMPUTE / RECODE로 적습니다(export-spss.js).
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  const DERIVE_TYPES = { compute: '값 계산', bin: '구간 나누기', merge: '보기 묶기' };
  const OPS = {
    'k-x': { label: '기준값 - 값', fn: (x, k) => k - x, sps: (v, k) => `${k}-${v}` },
    'x-k': { label: '값 - 기준값', fn: (x, k) => x - k, sps: (v, k) => `${v}-${k}` },
    'x*k': { label: '값 × 기준값', fn: (x, k) => x * k, sps: (v, k) => `${v}*${k}` },
    'x/k': { label: '값 ÷ 기준값', fn: (x, k) => (k === 0 ? null : x / k), sps: (v, k) => `${v}/${k}` },
  };

  // SPSS 변수 이름 규칙 (글자로 시작, 글자·숫자·_ 만)
  const NAME_RE = /^[A-Za-z가-힣][A-Za-z0-9가-힣_]*$/;

  const fmtNum = (n) => String(Math.round(n * 1e10) / 1e10);

  /**
   * "~4 / 5~9 / 10~" → [{lo:null,hi:4,label:'4 이하'}, {lo:5,hi:9,label:'5~9'}, {lo:10,hi:null,label:'10 이상'}]
   * 구분: / , 줄바꿈.  범위: "a~b"(a 이상 b 이하), "~b", "a~", 한 값 "a".
   */
  function parseBins(text) {
    const parts = String(text || '').split(/[\/,\n]+/).map((s) => s.trim()).filter(Boolean);
    if (!parts.length) return { error: '구간을 입력하세요. 예) ~4 / 5~9 / 10~' };
    const num = (s) => {
      s = s.trim();
      if (s === '') return null;
      const n = Number(s);
      return Number.isFinite(n) ? n : NaN;
    };
    const bins = [];
    for (const p of parts) {
      const m = /^(.*?)[~∼～](.*)$/.exec(p.replace(/\s+/g, ''));
      let lo;
      let hi;
      if (m) {
        lo = num(m[1]);
        hi = num(m[2]);
        if (lo == null && hi == null) return { error: `'${p}': 구간의 시작이나 끝 중 하나는 숫자여야 합니다.` };
      } else {
        lo = hi = num(p);
      }
      if (Number.isNaN(lo) || Number.isNaN(hi)) return { error: `'${p}': 숫자로 읽을 수 없습니다.` };
      if (lo != null && hi != null && lo > hi) return { error: `'${p}': 앞의 숫자가 뒤의 숫자보다 큽니다.` };
      const label = lo == null ? `${fmtNum(hi)} 이하` : hi == null ? `${fmtNum(lo)} 이상` : lo === hi ? fmtNum(lo) : `${fmtNum(lo)}~${fmtNum(hi)}`;
      bins.push({ lo, hi, label });
    }
    // 겹치는 구간 확인
    for (let i = 0; i < bins.length; i++) {
      for (let j = i + 1; j < bins.length; j++) {
        const a = bins[i];
        const b = bins[j];
        const aLo = a.lo == null ? -Infinity : a.lo;
        const aHi = a.hi == null ? Infinity : a.hi;
        const bLo = b.lo == null ? -Infinity : b.lo;
        const bHi = b.hi == null ? Infinity : b.hi;
        if (aLo <= bHi && bLo <= aHi) return { error: `'${a.label}'와 '${b.label}' 구간이 겹칩니다.` };
      }
    }
    return { bins };
  }

  function binOf(x, bins) {
    for (let i = 0; i < bins.length; i++) {
      const b = bins[i];
      if ((b.lo == null || x >= b.lo) && (b.hi == null || x <= b.hi)) return i + 1;
    }
    return null;
  }

  const isNines = (v) => v != null && TG.NINES.test(String(v));

  // 보기 묶기: 새 보기 이름이 같은 원래 보기끼리 한 보기로. 이름을 비우면 빈 값.
  function mergeCodes(def, srcCodes) {
    const names = [];
    const map = new Map();
    srcCodes.forEach((c) => {
      const name = String((def.groups || {})[c.code] == null ? '' : def.groups[c.code]).trim();
      if (!name) return;
      if (!names.includes(name)) names.push(name);
      map.set(c.code, names.indexOf(name) + 1);
    });
    return { codes: names.map((label, i) => ({ code: i + 1, label })), map };
  }

  // 원래 변수에서 가공 변수로 쓸 수 있는 보기 목록 (보기 묶기용)
  function sourceCodes(items, key) {
    const it = items.find((i) => !i.isGroup && i.keys[0] === key);
    return it && it.codes.length ? TG.allCodes(it) : [];
  }

  /** 입력 확인. 문제가 없으면 null, 있으면 안내 문장 */
  function checkDerived(def, data, items, others) {
    const name = String(def.name || '').trim();
    if (!name) return '새 변수 이름을 입력하세요.';
    if (!NAME_RE.test(name) || name.length > 64) return '새 변수 이름은 글자로 시작하고 글자·숫자·밑줄(_)만 쓸 수 있습니다. (SPSS 변수명 규칙)';
    const key = TG.keyOf(name);
    const col = data.columns.get(key);
    if ((col && !(col.derived && col.derived.id === def.id)) || (others || []).some((o) => o.id !== def.id && TG.keyOf(o.name) === key)) return `'${name}' 변수가 이미 있습니다. 다른 이름을 쓰세요.`;
    const src = data.columns.get(TG.keyOf(def.source || ''));
    if (!src) return '원래 변수를 고르세요.';
    if (def.type === 'compute' || def.type === 'bin') {
      if (src.isText) return '문자로 된 변수는 계산할 수 없습니다.';
    }
    if (def.type === 'compute') {
      if (!OPS[def.op]) return '계산 방법을 고르세요.';
      if (def.k === '' || def.k == null || !Number.isFinite(Number(def.k))) return '기준값을 숫자로 입력하세요.';
      if (def.op === 'x/k' && Number(def.k) === 0) return '0으로 나눌 수 없습니다.';
    }
    if (def.type === 'bin') {
      const p = parseBins(def.bins);
      if (p.error) return p.error;
    }
    if (def.type === 'merge') {
      const codes = sourceCodes(items, TG.keyOf(def.source));
      if (!codes.length) return '보기가 있는 문항만 묶을 수 있습니다.';
      if (!mergeCodes(def, codes).codes.length) return '새 보기 이름을 하나 이상 입력하세요.';
    }
    return null;
  }

  /** 가공 변수 하나의 값·보기 계산 */
  function derive(def, data, items) {
    const srcKey = TG.keyOf(def.source);
    const src = data.columns.get(srcKey);
    const n = data.n;
    const values = new Array(n).fill(null);
    let codes = [];
    let dropped = 0; // 어느 구간에도 들지 않은 값
    if (def.type === 'compute') {
      const k = Number(def.k);
      const fn = OPS[def.op].fn;
      for (let i = 0; i < n; i++) {
        const x = src.values[i];
        if (x == null || isNines(x)) continue;
        const y = fn(x, k);
        values[i] = y == null || !Number.isFinite(y) ? null : Math.round(y * 1e10) / 1e10;
      }
    } else if (def.type === 'bin') {
      const bins = parseBins(def.bins).bins;
      codes = bins.map((b, i) => ({ code: i + 1, label: b.label }));
      for (let i = 0; i < n; i++) {
        const x = src.values[i];
        if (x == null || isNines(x)) continue;
        values[i] = binOf(x, bins);
        if (values[i] == null) dropped++;
      }
    } else if (def.type === 'merge') {
      const m = mergeCodes(def, sourceCodes(items, srcKey));
      codes = m.codes;
      for (let i = 0; i < n; i++) {
        const x = src.values[i];
        if (x != null && m.map.has(x)) values[i] = m.map.get(x);
      }
    }
    return { values, codes, dropped, srcKey };
  }

  function describe(def) {
    if (def.type === 'compute') {
      const k = fmtNum(Number(def.k));
      return { 'k-x': `${k} - ${def.source}`, 'x-k': `${def.source} - ${k}`, 'x*k': `${def.source} × ${k}`, 'x/k': `${def.source} ÷ ${k}` }[def.op];
    }
    if (def.type === 'bin') return `${def.source} 구간: ${parseBins(def.bins).bins.map((b) => b.label).join(' / ')}`;
    // 예) CITY: 수도권(2개) / 그 외(14개)
    const cnt = new Map();
    Object.values(def.groups || {}).forEach((g) => {
      g = String(g == null ? '' : g).trim();
      if (g) cnt.set(g, (cnt.get(g) || 0) + 1);
    });
    return `${def.source}: ${Array.from(cnt).map(([g, c]) => `${g}(${c}개)`).join(' / ')}`;
  }

  /**
   * 가공 변수를 데이터·문항 목록에 반영합니다. 앞서 만든 가공 변수는 지우고 다시 만듭니다.
   * 문항 목록에서 고친 표 제목·포함·정렬 설정은 유지합니다.
   * 반환: { items, warnings }
   */
  function applyDerived(data, items, defs) {
    const prev = new Map();
    items.forEach((it) => {
      if (it.derived) prev.set(it.derived.id, it);
    });
    items = items.filter((it) => !it.derived);
    data.order = data.order.filter((k) => {
      const c = data.columns.get(k);
      if (c && c.derived) {
        data.columns.delete(k);
        return false;
      }
      return true;
    });
    const warnings = [];
    const done = [];
    (defs || []).forEach((def) => {
      const err = checkDerived(def, data, items, done);
      if (err) {
        warnings.push(`가공 변수 ${def.name || '(이름 없음)'}: ${err} → 만들지 않았습니다.`);
        return;
      }
      const r = derive(def, data, items);
      const key = TG.keyOf(def.name);
      const filled = r.values.filter((v) => v != null).length;
      data.columns.set(key, { name: def.name.trim(), values: r.values, raw: null, filled, nonNum: 0, isText: false, derived: def });
      data.order.push(key);
      if (r.dropped) warnings.push(`가공 변수 ${def.name}: 어느 구간에도 들지 않는 값 ${r.dropped}건은 빈 값으로 두었습니다.`);
      const old = prev.get(def.id);
      const label = String(def.label || '').trim() || def.name;
      const kind = def.type === 'compute' ? 'numeric' : 'single';
      items.push({
        id: 'd' + def.id,
        kind,
        autoKind: kind,
        reason: `가공: ${DERIVE_TYPES[def.type]}`,
        vars: [def.name.trim()],
        keys: [key],
        title: old && old.derivedLabel === label ? old.title : label,
        derivedLabel: label,
        itemLabels: [],
        codes: r.codes,
        extraCodes: [],
        include: old ? old.include : true,
        sort: old ? old.sort : false,
        isGroup: false,
        derived: def,
        meanOf: def.type === 'bin' ? r.srcKey : null,
      });
      done.push(def);
    });
    return { items, warnings };
  }

  Object.assign(TG, { DERIVE_TYPES, DERIVE_OPS: OPS, parseBins, previewDerived: derive, checkDerived, applyDerived, describeDerived: describe, sourceCodes, mergeCodes });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

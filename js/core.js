/*
 * 테이블 생성기 - 계산 엔진
 * 코드북 해석 → 문항 유형 자동 인식 → 표 계산
 * 브라우저(window.TG)와 Node(테스트용) 모두에서 동작합니다.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  // ------------------------------------------------------------------
  // 공통 유틸
  // ------------------------------------------------------------------
  function str(v) {
    return v == null ? '' : String(v).trim();
  }

  function toNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const s = String(v).trim();
    if (s === '') return null;
    const n = Number(s.replace(/,/g, ''));
    return isFinite(n) ? n : null;
  }

  function keyOf(name) {
    return str(name).toLowerCase();
  }

  // ------------------------------------------------------------------
  // 문항 유형
  // ------------------------------------------------------------------
  const KINDS = {
    single: '단일응답',
    scale: '척도',
    numeric: '수치(평균)',
    singleset: '단일응답 묶음',
    scaleset: '척도 묶음',
    multi01: '복수응답',
    rank: '순위응답',
    multicode: '복수응답(코드형)',
    exclude: '제외',
    weight: '가중치',
  };
  const SINGLE_KINDS = ['single', 'scale', 'numeric', 'exclude', 'weight'];
  const GROUP_KINDS = ['multi01', 'rank', 'multicode', 'scaleset', 'singleset', 'exclude'];

  function normType(s) {
    s = str(s).toLowerCase();
    if (!s) return '';
    if (/가중|weight/.test(s)) return 'weight';
    if (/순위|rank/.test(s)) return 'rank';
    if (/(복수|중복|multi).*(코드|code)|^mc$/.test(s)) return 'multicode';
    if (/복수|중복|multi|^ma$|0\s*\/\s*1|다중/.test(s)) return 'multi01';
    if (/척도|scale|리커트|likert/.test(s)) return 'scale';
    if (/단일|single|^sa$|범주/.test(s)) return 'single';
    if (/수치|숫자|연속|num|평균|정량/.test(s)) return 'numeric';
    if (/제외|개방|오픈|open|text|문자|^id$|skip|없음/.test(s)) return 'exclude';
    return '';
  }

  // ------------------------------------------------------------------
  // 코드북 해석
  //   입력: 2차원 배열(시트의 행들)
  //   지원 형식
  //     (1) 세로형: 변수명 | 문항 | 코드 | 레이블 | (유형) | (그룹)
  //         보기마다 1행, 변수명·문항은 첫 행에만 있어도 됨
  //     (2) 변수당 1행: 보기 칸에 "1=남자, 2=여자" 처럼 묶어서 작성
  // ------------------------------------------------------------------
  function classifyHeader(h) {
    const s = str(h).toLowerCase().replace(/\s+/g, '');
    if (!s) return '';
    if (/유형|type|형태|종류/.test(s)) return 'type';
    if (/그룹|group|묶음/.test(s)) return 'group';
    if (/변수(레이블|라벨|설명)|문항|질문|question|varlabel|variablelabel/.test(s)) return 'question';
    if (/(코드|code|^값$|^값\(|value$|^value|보기번호|번호$)/.test(s) && !/(레이블|라벨|label|명$)/.test(s)) return 'code';
    if (/변수|variable|^var|^name$/.test(s)) return 'var';
    if (/보기|레이블|라벨|label|응답|내용/.test(s)) return 'label';
    return '';
  }

  function findHeader(rows) {
    const limit = Math.min(rows.length, 20);
    for (let r = 0; r < limit; r++) {
      const row = rows[r] || [];
      const cols = {};
      row.forEach((cell, c) => {
        const k = classifyHeader(cell);
        if (k && cols[k] == null) cols[k] = c;
      });
      if (cols.var != null && Object.keys(cols).length >= 2) return { row: r, cols };
    }
    return null;
  }

  // "1=남자, 2=여자" / "1) 매우 그렇다 2) 그렇다" / "1:예 / 2:아니오" 형태 해석
  function parseInlineCodes(text) {
    const s = str(text);
    if (!s) return [];
    const re = /(^|[\s,;\/|\n])(-?\d+)\s*[=:)\.]\s*/g;
    const marks = [];
    let m;
    while ((m = re.exec(s))) {
      marks.push({ code: Number(m[2]), start: m.index + m[1].length, labelStart: re.lastIndex });
    }
    if (marks.length < 2) return [];
    return marks.map((mk, i) => {
      const end = i + 1 < marks.length ? marks[i + 1].start : s.length;
      const label = s.slice(mk.labelStart, end).replace(/[\s,;\/|]+$/, '').trim();
      return { code: mk.code, label };
    });
  }

  function parseCodebook(rows) {
    rows = (rows || []).map((r) => r || []);
    const header = findHeader(rows);
    const cols = header ? header.cols : { var: 0, question: 1, code: 2, label: 3, type: 4, group: 5 };
    const start = header ? header.row + 1 : 0;
    const vars = new Map();
    const warnings = [];
    let cur = null;
    const cell = (row, k) => (cols[k] == null ? '' : row[cols[k]]);

    for (let r = start; r < rows.length; r++) {
      const row = rows[r];
      const name = str(cell(row, 'var'));
      if (name) {
        const key = keyOf(name);
        if (!vars.has(key)) vars.set(key, { name, question: '', codes: [], type: '', group: '' });
        cur = vars.get(key);
      }
      if (!cur) continue;
      const q = str(cell(row, 'question'));
      if (q && !cur.question) cur.question = q;
      const t = normType(cell(row, 'type'));
      if (t && !cur.type) cur.type = t;
      const g = str(cell(row, 'group'));
      if (g && !cur.group) cur.group = g;

      const rawCode = cell(row, 'code');
      const label = str(cell(row, 'label'));
      const code = toNum(rawCode);
      if (code != null && str(rawCode) !== '') {
        if (!cur.codes.some((c) => c.code === code)) cur.codes.push({ code, label: label || String(code) });
      } else {
        const inline = parseInlineCodes(str(rawCode)).concat(parseInlineCodes(label));
        inline.forEach((c) => {
          if (!cur.codes.some((x) => x.code === c.code)) cur.codes.push(c);
        });
      }
    }
    if (!header) warnings.push('코드북 제목행(변수명/문항/코드/레이블)을 찾지 못해 A~F열 순서로 읽었습니다.');
    if (vars.size === 0) warnings.push('코드북에서 변수를 하나도 찾지 못했습니다.');
    return { vars, warnings, header: header ? header.cols : null };
  }

  // ------------------------------------------------------------------
  // 데이터 준비
  //   입력: 2차원 배열(1행 = 변수명)
  // ------------------------------------------------------------------
  function prepareData(rows) {
    rows = (rows || []).filter((r) => r && r.some((c) => str(c) !== ''));
    if (!rows.length) throw new Error('데이터가 비어 있습니다.');
    const names = rows[0].map((h) => str(h));
    const body = rows.slice(1);
    const columns = new Map();
    const order = [];
    names.forEach((name, c) => {
      if (!name) return;
      const values = new Array(body.length);
      const raw = new Array(body.length);
      let nonNum = 0;
      let filled = 0;
      for (let i = 0; i < body.length; i++) {
        const cell = body[i][c];
        const n = toNum(cell);
        values[i] = n;
        raw[i] = cell;
        if (str(cell) !== '') {
          filled++;
          if (n == null) nonNum++;
        }
      }
      const key = keyOf(name);
      if (columns.has(key)) return;
      const isText = nonNum > 0 && nonNum >= filled * 0.5;
      columns.set(key, { name, values, raw: isText ? raw : null, filled, nonNum, isText });
      order.push(key);
    });
    return { columns, order, n: body.length };
  }

  function distinctValues(values, limit) {
    const set = new Set();
    for (const v of values) {
      if (v == null) continue;
      set.add(v);
      if (limit && set.size > limit) break;
    }
    return set;
  }

  // ------------------------------------------------------------------
  // 자동 인식
  // ------------------------------------------------------------------
  const SCALE_WORD = /(전혀|매우|보통|그렇다|그렇지|만족|불만|동의|반대|중요|좋|나쁘|편리|불편|필요|높|낮|긍정|부정|찬성|많|적|약간|대체로|다소|도움|강함|약함|의향|쉽|어렵|\d+\s*점)/;
  const DK_WORD = /(모름|무응답|해당\s*없|잘\s*모|기타|거절|dk|n\/a)/i;
  const ID_NAME = /^(id|no|num|seq|번호|연번|일련|resp|sampleid|pid|uid|caseid|serial)(_?\d*)?$/i;
  const WEIGHT_NAME = /^(w|wt|wgt|weight|weights|가중치|가중)(_?\w*)?$/i;
  // 온라인 조사 시스템이 붙이는 관리용 열
  const SYSTEM_NAME = /^(group_id|start_time|end_time|loi(\(s\))?|loi_sum|access_key|user_agent.*|user_device|is_mobile|ip|data_time|result|last_question|last_before_question|status|duration|panel.*|device|browser)$/i;

  function scaleInfo(codes) {
    const valid = codes.filter((c) => !DK_WORD.test(c.label)).sort((a, b) => a.code - b.code);
    if (valid.length < 3 || valid.length > 11) return null;
    for (let i = 1; i < valid.length; i++) if (valid[i].code !== valid[i - 1].code + 1) return null;
    return valid;
  }

  function looksLikeScale(codes) {
    const valid = scaleInfo(codes);
    if (!valid) return false;
    const hits = valid.filter((c) => SCALE_WORD.test(c.label)).length;
    return hits >= Math.ceil(valid.length / 2);
  }

  // "<--", "---" 처럼 기호만 있는 척도 보기 이름을 "3점"처럼 바꿉니다.
  function fixLabels(codes) {
    const placeholder = (l) => !str(l) || /^[\s<>\-=~.·→←]+$/.test(l);
    if (!codes.some((c) => placeholder(c.label))) return codes;
    return codes.map((c) => ({ code: c.code, label: placeholder(c.label) ? `${c.code}점` : /^\d+\s*점/.test(c.label) ? c.label : `${c.code}점(${str(c.label).replace(/\s+/g, ' ')})` }));
  }

  function sameCodes(a, b) {
    if (a.length !== b.length) return false;
    return a.every((c, i) => c.code === b[i].code);
  }

  function commonPrefix(list) {
    if (!list.length) return '';
    let p = list[0];
    for (const s of list) {
      let i = 0;
      while (i < p.length && i < s.length && p[i] === s[i]) i++;
      p = p.slice(0, i);
    }
    // 단어 중간에서 잘리면("사업A/사업B" → "A/B") 앞의 구분 기호까지 물러남
    const word = /[0-9A-Za-z가-힣]/;
    if (p && word.test(p[p.length - 1]) && list.some((s) => s.length > p.length && word.test(s[p.length]))) {
      let j = p.length - 1;
      while (j >= 0 && word.test(p[j])) j--;
      p = p.slice(0, j + 1);
    }
    return p;
  }

  function cleanTitle(s) {
    return str(s)
      .replace(/^\[[^\]]*순위\]\s*/, '')
      .replace(/[\s\-–—:·_(（\[▶◆■※]+$/, '')
      .trim();
  }

  function groupLabels(vars, cbOf, kind, data) {
    const cbs = vars.map((v) => cbOf(v));
    const qs = cbs.map((c) => str(c.question));
    const allHave = qs.every((q) => q);
    const groupTitle = cbs.map((c) => c.groupTitle).find(Boolean);
    let title = '';
    let items;
    if (cbs.every((c) => c.itemLabel)) {
      title = groupTitle || cleanTitle(commonPrefix(qs)) || qs[0];
      items = cbs.map((c) => c.itemLabel);
    } else if (allHave && new Set(qs).size > 1) {
      const pre = commonPrefix(qs);
      title = cleanTitle(pre) || qs[0];
      items = qs.map((q) => str(q.slice(pre.length)).replace(/^[\s\-–—:·_)）\]▶]+/, '').replace(/^([^(（]*)[)）]$/, '$1').trim() || q);
    } else {
      title = qs.find((q) => q) || vars[0];
      items = vars.map((v, i) => {
        if (kind === 'rank') return i + 1 + '순위';
        const cb = cbOf(v);
        if (kind === 'multi01') {
          // 선택 시 보기 번호가 들어가는 방식이면 그 번호의 보기 이름
          const vals = Array.from(distinctValues(data.columns.get(v).values, 5)).filter((x) => x !== 0);
          if (vals.length === 1) {
            const hit = cb.codes.find((c) => c.code === vals[0]);
            if (hit && !/^(예|선택|해당|yes|있음|체크)$/i.test(hit.label)) return hit.label;
          }
          const c1 = cb.codes.find((c) => c.code === 1);
          if (c1 && cb.codes.length <= 2 && !/^(예|선택|해당|yes|1|있음|체크)$/i.test(c1.label)) return c1.label;
          const m = /(\d+)$/.exec(cb.name || v);
          const byPos = m && cb.codes.find((c) => c.code === Number(m[1]));
          if (byPos && cb.codes.length > 2) return byPos.label;
        }
        return cb.name || v;
      });
    }
    return { title: cleanTitle(groupTitle || title) || title, items };
  }

  function shortName(title, fallback) {
    let s = str(title)
      .replace(/^\[?[A-Za-z]{0,4}\d+[\w\-]*[.)\]]?\s*/, '')
      .replace(/\s*▶\s*/g, ' ')
      .replace(/[?？].*$/, '')
      .replace(/^(귀하의|귀하는|귀하께서는|귀하께서|귀댁의|귀댁은|귀사의|귀사는|현재)\s*/, '')
      .replace(/\s*(은|는|이|가)?\s*(무엇|어떻게|어디|몇|얼마).*$/, '')
      .trim();
    if (!s) s = fallback;
    return s.length > 20 ? s.slice(0, 20) + '…' : s;
  }

  // 코드북에 없는 숫자 변수의 보기 이름을, 1:1로 대응하는 문자 열에서 찾아옵니다.
  // 예) DE1(1,2,3) ↔ 주체분류("01.대기업","02.중견기업","03.중소기업")
  function findLabelColumn(key, data) {
    const col = data.columns.get(key);
    const d = distinctValues(col.values, 40);
    if (d.size < 2 || d.size > 40 || !Array.from(d).every((v) => Number.isInteger(v))) return null;
    let best = null;
    for (const k2 of data.order) {
      const other = data.columns.get(k2);
      if (k2 === key || !other.isText || !other.raw) continue;
      const fwd = new Map();
      const back = new Map();
      let ok = true;
      let pairs = 0;
      for (let i = 0; i < data.n && ok; i++) {
        const v = col.values[i];
        const t = str(other.raw[i]);
        if (v == null || !t) {
          if ((v == null) !== !t) ok = false;
          continue;
        }
        pairs++;
        if (fwd.has(v) && fwd.get(v) !== t) ok = false;
        if (back.has(t) && back.get(t) !== v) ok = false;
        fwd.set(v, t);
        back.set(t, v);
      }
      if (!ok || pairs < Math.max(2, col.filled * 0.95)) continue;
      const numbered = Array.from(fwd.entries()).every(([v, t]) => new RegExp('^0*' + v + '\\s*[.)]').test(t));
      const codes = Array.from(fwd.entries())
        .sort((a, b) => a[0] - b[0])
        .map(([v, t]) => ({ code: v, label: t.replace(/^\d+\s*[.)]\s*/, '') }));
      const cand = { source: other.name, codes, score: numbered ? 2 : 1 };
      if (!best || cand.score > best.score) best = cand;
    }
    return best;
  }

  /**
   * 코드북과 데이터를 비교해 표로 만들 "문항 목록"을 자동 생성합니다.
   */
  function buildItems(codebook, data) {
    const warnings = [];
    const cbVars = codebook ? codebook.vars : new Map();
    const hasCodebook = cbVars.size > 0;
    const blank = new Map();
    const cbOf = (key) => {
      if (cbVars.has(key)) return cbVars.get(key);
      if (!blank.has(key)) blank.set(key, { name: data.columns.get(key) ? data.columns.get(key).name : key, question: '', codes: [], type: '', group: '' });
      return blank.get(key);
    };

    // 코드북에만 있고 데이터에 없는 변수
    const missing = [];
    cbVars.forEach((v, key) => {
      if (!data.columns.has(key)) missing.push(v.name);
    });
    if (missing.length) warnings.push(`코드북에는 있지만 데이터에 없는 변수 ${missing.length}개: ${missing.slice(0, 15).join(', ')}${missing.length > 15 ? ' …' : ''}`);

    // 1) 그룹 후보 묶기
    const groupOf = new Map();
    const groups = new Map();
    data.order.forEach((key) => {
      const cb = cbOf(key);
      let g = cb.group ? 'G:' + keyOf(cb.group) : '';
      if (!g) {
        const m = /^(.+?)[_.\-](\d+)$/.exec(data.columns.get(key).name);
        if (m) g = 'P:' + keyOf(m[1]);
      }
      if (!g) return;
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(key);
      groupOf.set(key, g);
    });
    // 이름 패턴으로 묶였지만 개별 문항으로 봐야 하는 경우 풀기
    groups.forEach((keys, g) => {
      let split = keys.length < 2;
      if (!split && g.startsWith('P:')) {
        const types = keys.map((k) => cbOf(k).type);
        const explicitSingle = types.every((t) => ['single', 'scale', 'numeric', 'weight', 'exclude'].includes(t));
        // 보기 없는 연속형 숫자(종사자 수, 금액 등)
        const numericLike = keys.every((k) => {
          if (cbOf(k).codes.length) return false;
          const col = data.columns.get(k);
          if (col.isText) return true;
          const d = distinctValues(col.values, 60);
          return d.size > 50 || Array.from(d).some((v) => !Number.isInteger(v));
        });
        const anyText = keys.some((k) => data.columns.get(k).isText);
        split = explicitSingle || numericLike || anyText;
      }
      if (split) {
        keys.forEach((k) => groupOf.delete(k));
        groups.delete(g);
      }
    });

    const items = [];
    const done = new Set();
    let seq = 0;

    function codesFor(keys) {
      let best = [];
      for (const k of keys) if (cbOf(k).codes.length > best.length) best = cbOf(k).codes;
      return fixLabels(best.slice());
    }

    function addUnknownCodes(codes, keys) {
      const known = new Set(codes.map((c) => c.code));
      const extra = new Set();
      keys.forEach((k) => distinctValues(data.columns.get(k).values, 200).forEach((v) => {
        if (!known.has(v)) extra.add(v);
      }));
      return Array.from(extra).sort((a, b) => a - b);
    }

    data.order.forEach((key) => {
      if (done.has(key)) return;
      const col = data.columns.get(key);
      const g = groupOf.get(key);

      // ---------- 묶음 문항 ----------
      if (g) {
        const keys = groups.get(g);
        keys.forEach((k) => done.add(k));
        const explicit = keys.map((k) => cbOf(k).type).find((t) => t && t !== 'single' && t !== 'numeric') || '';
        const codes = codesFor(keys);
        let kind = '';
        let reason = '';
        const perVar = keys.map((k) => distinctValues(data.columns.get(k).values, 50));
        const allVals = new Set();
        perVar.forEach((d) => d.forEach((v) => allVals.add(v)));
        const is01 = allVals.size > 0 && Array.from(allVals).every((v) => v === 0 || v === 1) && codes.every((c) => c.code === 0 || c.code === 1);
        // 선택하면 보기 번호, 안 하면 빈칸(또는 0): Q_1=1, Q_2=2, Q_3=3 …
        const nonZero = perVar.map((d) => Array.from(d).filter((v) => v !== 0));
        const positional = !is01 && nonZero.every((d) => d.length <= 1) && new Set(nonZero.flat()).size === nonZero.filter((d) => d.length).length && nonZero.filter((d) => d.length).length >= 2;
        const textAll = keys.map((k) => cbOf(k).question).join(' ');

        if (explicit === 'exclude') {
          kind = 'exclude'; reason = '코드북 지정';
        } else if (explicit === 'scale') {
          kind = 'scaleset'; reason = '코드북 지정';
        } else if (['multi01', 'rank', 'multicode'].includes(explicit)) {
          kind = explicit; reason = '코드북 지정';
        } else if (is01) {
          kind = 'multi01'; reason = '값이 0/1로만 구성';
        } else if (positional) {
          kind = 'multi01'; reason = '선택한 보기 번호/빈칸으로 구성';
        } else {
          const codeSets = keys.map((k) => cbOf(k).codes).filter((c) => c.length);
          const sorted = (c) => c.slice().sort((a, b) => a.code - b.code);
          const shared = codeSets.length > 0 && codeSets.every((c) => sameCodes(sorted(c), sorted(codeSets[0])));
          if (shared && looksLikeScale(codes)) {
            kind = 'scaleset'; reason = '같은 척도 보기 공유';
          } else if (shared || codeSets.length <= 1) {
            // 한 응답자 안에서 값이 중복되지 않으면 순위/복수(코드형)
            let rowsWithAny = 0;
            let dup = 0;
            for (let i = 0; i < data.n; i++) {
              const seen = new Set();
              let any = false;
              for (const k of keys) {
                const v = data.columns.get(k).values[i];
                if (v == null) continue;
                any = true;
                if (seen.has(v)) { dup++; break; }
                seen.add(v);
              }
              if (any) rowsWithAny++;
            }
            const noDup = rowsWithAny > 0 && dup / rowsWithAny < 0.05;
            if (noDup && (codes.length >= keys.length || codes.length === 0)) {
              if (/순위|순서|우선/.test(textAll) || !/(모두|복수|중복|해당하는\s*것)/.test(textAll)) {
                kind = 'rank'; reason = '같은 보기 · 응답자별 중복 없음';
              } else {
                kind = 'multicode'; reason = '문항에 "모두/복수" 표현';
              }
            } else {
              kind = 'singleset'; reason = '같은 보기를 쓰는 개별 문항';
            }
          } else {
            kind = 'singleset'; reason = '보기가 서로 다른 개별 문항';
          }
        }
        if (!codes.length && (kind === 'rank' || kind === 'multicode')) {
          warnings.push(`${col.name} 등 ${keys.length}개 변수: 코드북에 보기가 없어 코드 번호로 표시합니다.`);
        }
        const lab = groupLabels(keys, cbOf, kind, data);
        items.push({
          id: 'i' + ++seq,
          kind,
          autoKind: kind,
          reason,
          vars: keys.map((k) => data.columns.get(k).name),
          keys,
          title: lab.title,
          itemLabels: lab.items,
          codes,
          extraCodes: kind === 'multi01' ? [] : addUnknownCodes(codes, keys),
          include: kind !== 'exclude',
          isGroup: true,
        });
        return;
      }

      // ---------- 개별 문항 ----------
      done.add(key);
      const cb = cbOf(key);
      const inCodebook = cbVars.has(key);
      let codes = fixLabels(cb.codes.slice());
      let kind = cb.type && SINGLE_KINDS.includes(cb.type) ? cb.type : '';
      let reason = kind ? '코드북 지정' : '';
      let include = true;
      if (kind === 'scale' && !scaleInfo(codes)) kind = codes.length ? 'single' : 'numeric';
      if (!kind && cb.type && cb.type !== 'single') {
        kind = cb.type === 'exclude' ? 'exclude' : ''; // 묶음 유형이 개별 변수에 지정된 경우 무시
        if (kind) reason = '코드북 지정';
      }
      if (!kind) {
        const dAll = distinctValues(col.values);
        const allInt = Array.from(dAll).every((v) => Number.isInteger(v));
        if (WEIGHT_NAME.test(col.name)) {
          kind = 'weight'; reason = '변수명이 가중치 형태';
        } else if (SYSTEM_NAME.test(col.name)) {
          kind = 'exclude'; reason = '조사 시스템 관리 변수';
        } else if (col.isText) {
          kind = 'exclude'; reason = '문자(개방형) 응답';
        } else if (ID_NAME.test(col.name)) {
          kind = 'exclude'; reason = 'ID 변수';
        } else if (codes.length) {
          if (looksLikeScale(codes)) { kind = 'scale'; reason = '척도형 보기'; }
          else { kind = 'single'; reason = '코드북 보기 있음'; }
        } else if (dAll.size === 0) {
          kind = 'exclude'; reason = '응답 없음';
        } else if (dAll.size >= Math.max(20, data.n * 0.9) && allInt) {
          kind = 'exclude'; reason = '값이 모두 달라 ID로 추정';
        } else {
          const found = !inCodebook ? findLabelColumn(key, data) : null;
          if (found) {
            kind = 'single'; reason = `보기 이름을 '${found.source}' 열에서 가져옴`;
            codes = found.codes;
            const srcCb = cbVars.get(keyOf(found.source));
            cb.question = cb.question || (srcCb && srcCb.question) || found.source;
          } else if (!inCodebook && !hasCodebook && allInt && dAll.size <= 10) {
            kind = 'single'; reason = '코드북에 없음 · 정수 코드';
            codes = Array.from(dAll).sort((a, b) => a - b).map((v) => ({ code: v, label: '코드 ' + v }));
            warnings.push(`${col.name}: 코드북에 없어 코드 값을 그대로 보기로 사용합니다.`);
          } else { kind = 'numeric'; reason = '보기 없는 숫자'; }
        }
        // 코드북이 있는데 코드북에 없는 변수(표본 정보 등)는 기본으로 표를 만들지 않음 (배너로는 사용 가능)
        if (hasCodebook && !inCodebook && kind !== 'exclude' && kind !== 'weight') {
          include = false;
          reason += ' · 코드북에 없는 변수';
        }
      }
      items.push({
        id: 'i' + ++seq,
        kind,
        autoKind: kind,
        reason,
        vars: [col.name],
        keys: [key],
        title: cb.question || col.name,
        itemLabels: [],
        codes,
        extraCodes: ['single', 'scale'].includes(kind) ? addUnknownCodes(codes, [key]) : [],
        include: include && !['exclude', 'weight'].includes(kind),
        isGroup: false,
      });
    });

    // 같은 제목의 표가 여러 개면 변수명을 붙여 구분
    const seen = new Map();
    items.forEach((it) => {
      if (!it.include) return;
      seen.set(it.title, (seen.get(it.title) || 0) + 1);
    });
    items.forEach((it) => {
      if (it.include && seen.get(it.title) > 1) {
        const tag = it.isGroup ? it.vars[0].replace(/[_.\-]\d+$/, '') : it.vars[0];
        it.title = `${it.title} [${tag}]`;
      }
    });

    items.forEach((it) => {
      if (it.extraCodes.length && it.include) {
        warnings.push(`${it.vars.join(', ')}: 코드북에 없는 값 발견 (${it.extraCodes.slice(0, 10).join(', ')}) → "미정의 코드"로 표시합니다.`);
      }
    });
    return { items, warnings };
  }

  // ------------------------------------------------------------------
  // 표 계산
  // ------------------------------------------------------------------
  function allCodes(item) {
    const list = item.codes.map((c) => ({ code: c.code, label: c.label }));
    (item.extraCodes || []).forEach((v) => list.push({ code: v, label: `미정의 코드(${v})`, undefinedCode: true }));
    return list;
  }

  function buildSegments(data, items, opts) {
    const n = data.n;
    const all = new Array(n).fill(true);
    const segs = [{ group: '전체', label: '전체', mask: all, isTotal: true }];
    (opts.banners || []).forEach((b) => {
      const key = keyOf(b.var);
      const col = data.columns.get(key);
      if (!col) return;
      const item = items.find((it) => it.keys.length === 1 && it.keys[0] === key);
      const codes = b.codes ? b.codes : item ? allCodes(item) : Array.from(distinctValues(col.values)).sort((a, c) => a - c).map((v) => ({ code: v, label: String(v) }));
      codes.forEach((c) => {
        const mask = col.values.map((v) => v === c.code);
        if (!mask.some(Boolean) && c.undefinedCode) return;
        segs.push({ group: b.label || (item ? shortName(item.title, col.name) : col.name), label: c.label, mask });
      });
    });
    return segs;
  }

  function weightsOf(data, opts) {
    const w = new Array(data.n).fill(1);
    if (!opts.weightVar) return w;
    const col = data.columns.get(keyOf(opts.weightVar));
    if (!col) return w;
    for (let i = 0; i < data.n; i++) {
      const v = col.values[i];
      w[i] = v == null || v < 0 ? 0 : v;
    }
    return w;
  }

  function pct(a, b) {
    return b > 0 ? (a / b) * 100 : null;
  }

  function colVals(data, key) {
    return data.columns.get(key).values;
  }

  // 단일/척도 1개 변수의 분포 계산
  function distTable(vals, codes, segs, w, scale) {
    return segs.map((seg) => {
      let n = 0;
      let wn = 0;
      const cnt = new Map(codes.map((c) => [c.code, 0]));
      for (let i = 0; i < vals.length; i++) {
        if (!seg.mask[i]) continue;
        const v = vals[i];
        if (v == null) continue;
        n++;
        wn += w[i];
        if (cnt.has(v)) cnt.set(v, cnt.get(v) + w[i]);
      }
      const row = { group: seg.group, label: seg.label, isTotal: !!seg.isTotal, n, wn, values: codes.map((c) => pct(cnt.get(c.code), wn)) };
      if (scale) {
        let sw = 0;
        let sx = 0;
        let top = 0;
        let bot = 0;
        for (let i = 0; i < vals.length; i++) {
          if (!seg.mask[i]) continue;
          const v = vals[i];
          if (v == null || !scale.validSet.has(v)) continue;
          const x = scale.positiveHigh ? v : scale.min + scale.max - v;
          sw += w[i];
          sx += w[i] * x;
          if (scale.topSet.has(v)) top += w[i];
          if (scale.botSet.has(v)) bot += w[i];
        }
        const mean = sw > 0 ? sx / sw : null;
        row.scale = {
          top: pct(top, wn),
          bot: pct(bot, wn),
          mean,
          score: mean == null ? null : ((mean - scale.min) / (scale.max - scale.min)) * 100,
        };
      }
      return row;
    });
  }

  function scaleSpec(codes) {
    const valid = scaleInfo(codes);
    if (!valid) return null;
    const min = valid[0].code;
    const max = valid[valid.length - 1].code;
    const lowLab = valid[0].label;
    const highLab = valid[valid.length - 1].label;
    const posLow = /(매우|아주|항상).{0,3}(그렇다|만족|동의|중요|좋|필요|찬성|높|편리)/.test(lowLab) && !/않|불|없/.test(lowLab);
    const negHigh = /(전혀|매우\s*(불|나쁘|낮|반대)|아주\s*(불|나쁘))/.test(highLab);
    const positiveHigh = !(posLow || negHigh);
    const k = valid.length >= 4 ? 2 : 1;
    const ordered = valid.map((c) => c.code); // 오름차순
    const topCodes = positiveHigh ? ordered.slice(-k) : ordered.slice(0, k);
    const botCodes = positiveHigh ? ordered.slice(0, k) : ordered.slice(-k);
    return {
      min, max, k, positiveHigh,
      points: valid.length,
      validSet: new Set(ordered),
      topSet: new Set(topCodes),
      botSet: new Set(botCodes),
    };
  }

  function mkTable(item, title, kind, columns, rows, notes) {
    return { itemId: item.id, title, kind, columns, rows, notes: notes || [], vars: item.vars };
  }

  function baseNotes(opts, extra) {
    const notes = [];
    if (extra) notes.push(extra);
    if (opts.weightVar) notes.push(`가중치 적용(${opts.weightVar})`);
    return notes;
  }

  function singleTables(item, key, title, data, segs, w, opts, forceScale) {
    const codes = allCodes(item);
    const vals = colVals(data, key);
    const spec = forceScale ? scaleSpec(item.codes) : null;
    const rows = distTable(vals, codes, segs, w, spec);
    const columns = codes.map((c) => ({ label: c.label, fmt: 'pct' }));
    const notes = baseNotes(opts, 'Base: 해당 문항 응답자');
    if (opts.showSum !== false) {
      columns.push({ label: '계', fmt: 'pct', isSum: true });
      rows.forEach((r) => r.values.push(r.wn > 0 ? r.values.reduce((a, b) => a + (b || 0), 0) : null));
    }
    if (spec) {
      const kk = spec.k === 2 ? '2' : '1';
      columns.push({ label: `긍정(Top${kk})`, fmt: 'pct', isStat: true });
      columns.push({ label: `부정(Bottom${kk})`, fmt: 'pct', isStat: true });
      columns.push({ label: '평균', fmt: 'mean', isStat: true });
      columns.push({ label: '100점 환산', fmt: 'mean1', isStat: true });
      rows.forEach((r) => r.values.push(r.scale.top, r.scale.bot, r.scale.mean, r.scale.score));
      notes.push(`${spec.points}점 척도 평균(${spec.min}~${spec.max}점)${spec.positiveHigh ? '' : ', 긍정일수록 높은 점수가 되도록 역산'}; 100점 환산 = (평균-${spec.min})÷${spec.max - spec.min}×100`);
      if (codes.length > spec.points) notes.push('모름/무응답 등은 평균 계산에서 제외');
    } else if (forceScale) {
      notes.push('보기가 연속된 척도 형태가 아니어서 평균을 계산하지 않았습니다.');
    }
    return { table: mkTable(item, title, forceScale ? 'scale' : 'single', columns, rows, notes), spec, rows };
  }

  function numericTable(item, data, segs, w, opts) {
    const vals = colVals(data, item.keys[0]);
    const rows = segs.map((seg) => {
      const xs = [];
      let n = 0;
      let sw = 0;
      let sx = 0;
      for (let i = 0; i < vals.length; i++) {
        if (!seg.mask[i] || vals[i] == null) continue;
        n++;
        sw += w[i];
        sx += w[i] * vals[i];
        xs.push([vals[i], w[i]]);
      }
      if (!n || sw <= 0) return { group: seg.group, label: seg.label, isTotal: !!seg.isTotal, n, wn: sw, values: [null, null, null, null, null] };
      const mean = sx / sw;
      let ss = 0;
      xs.forEach(([x, wi]) => (ss += wi * (x - mean) * (x - mean)));
      const sd = n > 1 ? Math.sqrt(ss / sw * (n / (n - 1))) : 0;
      xs.sort((a, b) => a[0] - b[0]);
      let acc = 0;
      let med = xs[xs.length - 1][0];
      for (const [x, wi] of xs) {
        acc += wi;
        if (acc >= sw / 2) { med = x; break; }
      }
      return { group: seg.group, label: seg.label, isTotal: !!seg.isTotal, n, wn: sw, values: [mean, sd, xs[0][0], med, xs[xs.length - 1][0]] };
    });
    const columns = ['평균', '표준편차', '최솟값', '중앙값', '최댓값'].map((l) => ({ label: l, fmt: 'mean' }));
    return mkTable(item, item.title, 'numeric', columns, rows, baseNotes(opts, 'Base: 해당 문항 응답자'));
  }

  // 여러 변수 중 하나라도 해당 코드를 가진 응답자 비율 (순위 종합/복수 코드형)
  function anyOfTable(item, keys, baseKeys, title, data, segs, w, opts, note) {
    const codes = allCodes(item);
    const cols = keys.map((k) => colVals(data, k));
    const baseCols = baseKeys.map((k) => colVals(data, k));
    const rows = segs.map((seg) => {
      let n = 0;
      let wn = 0;
      const cnt = new Map(codes.map((c) => [c.code, 0]));
      for (let i = 0; i < data.n; i++) {
        if (!seg.mask[i]) continue;
        if (!baseCols.some((c) => c[i] != null)) continue;
        n++;
        wn += w[i];
        const picked = new Set();
        cols.forEach((c) => {
          if (c[i] != null) picked.add(c[i]);
        });
        picked.forEach((v) => {
          if (cnt.has(v)) cnt.set(v, cnt.get(v) + w[i]);
        });
      }
      return { group: seg.group, label: seg.label, isTotal: !!seg.isTotal, n, wn, values: codes.map((c) => pct(cnt.get(c.code), wn)) };
    });
    const columns = codes.map((c) => ({ label: c.label, fmt: 'pct' }));
    return mkTable(item, title, 'multi', columns, rows, baseNotes(opts, note));
  }

  function multi01Table(item, data, segs, w, opts) {
    const cols = item.keys.map((k) => colVals(data, k));
    const rows = segs.map((seg) => {
      let n = 0;
      let wn = 0;
      const cnt = new Array(cols.length).fill(0);
      for (let i = 0; i < data.n; i++) {
        if (!seg.mask[i]) continue;
        // 0/1 방식은 0도 응답, 번호/빈칸 방식은 하나라도 선택한 사람이 Base
        if (!cols.some((c) => c[i] != null)) continue;
        n++;
        wn += w[i];
        cols.forEach((c, j) => {
          if (c[i] != null && c[i] !== 0) cnt[j] += w[i];
        });
      }
      return { group: seg.group, label: seg.label, isTotal: !!seg.isTotal, n, wn, values: cnt.map((c) => pct(c, wn)) };
    });
    const columns = item.itemLabels.map((l) => ({ label: l, fmt: 'pct' }));
    return mkTable(item, item.title, 'multi', columns, rows, baseNotes(opts, 'Base: 해당 문항 응답자, 복수응답(합계 100% 초과 가능)'));
  }

  function summaryTable(item, title, labels, rowsList, pick, fmt, opts, note) {
    const differ = rowsList.some((rs) => rs[0].n !== rowsList[0][0].n);
    const rows = rowsList[0].map((r, s) => {
      // 항목마다 Base가 다르면 가장 큰 Base를 표시
      const top = rowsList.map((rs) => rs[s]).reduce((a, b) => (b.n > a.n ? b : a));
      return { group: r.group, label: r.label, isTotal: r.isTotal, n: top.n, wn: top.wn, values: rowsList.map((rs) => pick(rs[s])) };
    });
    const columns = labels.map((l) => ({ label: l, fmt }));
    const notes = baseNotes(opts, note);
    if (differ) notes.push('항목마다 응답자(Base)가 달라 사례수는 가장 큰 값을 표시했습니다. 항목별 사례수는 개별 표를 참고하세요.');
    return mkTable(item, title, 'summary', columns, rows, notes);
  }

  /**
   * 선택된 문항과 옵션으로 모든 표를 계산합니다.
   * opts = { banners:[{var,label}], weightVar, decimals, showSum }
   */
  function computeTables(items, data, opts) {
    opts = opts || {};
    const segs = buildSegments(data, items, opts);
    const w = weightsOf(data, opts);
    const tables = [];
    items.forEach((item) => {
      if (!item.include) return;
      const kind = item.kind;
      try {
        if (kind === 'single' || kind === 'scale') {
          tables.push(singleTables(item, item.keys[0], item.title, data, segs, w, opts, kind === 'scale').table);
        } else if (kind === 'numeric') {
          tables.push(numericTable(item, data, segs, w, opts));
        } else if (kind === 'singleset' || kind === 'scaleset') {
          const rowsList = [];
          let spec = null;
          item.keys.forEach((k, j) => {
            const r = singleTables(item, k, `${item.title} - ${item.itemLabels[j]}`, data, segs, w, opts, kind === 'scaleset');
            tables.push(r.table);
            rowsList.push(r.rows);
            spec = spec || r.spec;
          });
          const sharedCodes = kind === 'singleset' && item.codes.length >= 2 && item.codes.length <= 3 &&
            item.keys.every((k) => { const v = data.columns.get(k).values; return v.every((x) => x == null || item.codes.some((c) => c.code === x)); });
          if (sharedCodes) {
            const first = item.codes[0];
            tables.push(summaryTable(item, `${item.title} - 요약: '${first.label}' 비율`, item.itemLabels, rowsList, (r) => r.values[0], 'pct', opts, `항목별 '${first.label}' 응답 비율(%)`));
          }
          if (kind === 'scaleset' && spec) {
            const kk = spec.k === 2 ? '2' : '1';
            tables.push(summaryTable(item, `${item.title} - 요약: 평균`, item.itemLabels, rowsList, (r) => r.scale.mean, 'mean', opts, `${spec.points}점 척도 평균`));
            tables.push(summaryTable(item, `${item.title} - 요약: 100점 환산`, item.itemLabels, rowsList, (r) => r.scale.score, 'mean1', opts, '100점 환산 점수'));
            tables.push(summaryTable(item, `${item.title} - 요약: 긍정(Top${kk}) 비율`, item.itemLabels, rowsList, (r) => r.scale.top, 'pct', opts, `긍정 응답(Top${kk}) 비율(%)`));
          }
        } else if (kind === 'multi01') {
          tables.push(multi01Table(item, data, segs, w, opts));
        } else if (kind === 'rank') {
          const first = singleTables(item, item.keys[0], `${item.title} - 1순위`, data, segs, w, { ...opts, showSum: opts.showSum }, false).table;
          first.kind = 'rank';
          tables.push(first);
          const k = item.keys.length;
          const label = k === 2 ? '1+2순위' : `1~${k}순위 종합`;
          tables.push(anyOfTable(item, item.keys, [item.keys[0]], `${item.title} - ${label}`, data, segs, w, opts, `Base: 1순위 응답자, ${label} 중복 응답(합계 100% 초과)`));
        } else if (kind === 'multicode') {
          tables.push(anyOfTable(item, item.keys, item.keys, item.title, data, segs, w, opts, 'Base: 해당 문항 응답자, 복수응답(합계 100% 초과 가능)'));
        }
      } catch (e) {
        tables.push({ itemId: item.id, title: item.title, kind: 'error', columns: [], rows: [], notes: ['계산 오류: ' + e.message], vars: item.vars });
      }
    });
    tables.forEach((t, i) => (t.no = i + 1));
    return { tables, n: data.n, segments: segs.map((s) => ({ group: s.group, label: s.label })) };
  }

  Object.assign(TG, {
    KINDS, SINGLE_KINDS, GROUP_KINDS,
    str, toNum, keyOf, normType, shortName,
    parseInlineCodes, parseCodebook, prepareData, buildItems, computeTables,
  });

  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

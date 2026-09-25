/*
 * 코드북 형식(시스템)별 읽기 규칙
 *
 * 조사 시스템마다 코드북 모양이 달라서, 형식마다 "어댑터"를 하나씩 둡니다.
 * 어댑터는 코드북 파일(모든 시트)을 받아 아래 공통 모양으로 바꿔 줍니다.
 *
 *   vars: Map(소문자 변수명 → {
 *     name, question, codes:[{code,label}],
 *     type,        // single / scale / multi01 / rank / numeric / exclude … ('' = 자동 인식)
 *     group,       // 같은 문항으로 묶을 이름 (없으면 변수명 규칙으로 자동 묶음)
 *     groupTitle,  // 묶음 문항 제목
 *     itemLabel,   // 묶음 안에서 이 변수의 항목 이름
 *   })
 *
 * 새 형식을 추가하려면 TG.registerFormat({...})로 등록하면 됩니다.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});
  const { str, toNum, keyOf } = TG;

  const FORMATS = [];
  function registerFormat(fmt) {
    FORMATS.push(fmt);
  }

  function findSheet(sheets, re) {
    return Object.keys(sheets).find((n) => re.test(n.trim()));
  }

  function headerIndex(row, re) {
    return (row || []).findIndex((c) => re.test(str(c)));
  }

  function newVar(name) {
    return { name, question: '', codes: [], type: '', group: '', groupTitle: '', itemLabel: '' };
  }

  // ------------------------------------------------------------------
  // 1) 표준 양식 (변수명 | 문항 | 코드 | 레이블 | 유형 | 그룹)
  // ------------------------------------------------------------------
  registerFormat({
    id: 'standard',
    name: '표준 양식(세로형)',
    usesSheet: true,
    detect: () => 1, // 다른 형식이 아니면 이 형식
    parse(sheets, opts) {
      const names = Object.keys(sheets);
      const sheet = (opts && opts.sheet && sheets[opts.sheet] ? opts.sheet : null) || names.find((n) => /코드|code/i.test(n)) || names[0];
      return TG.parseCodebook(sheets[sheet]);
    },
  });

  // ------------------------------------------------------------------
  // 2) INT64 (외주 조사 시스템)
  //    시트: GUIDE(문항유형) / GUIDE2(변수별 문항·보기) / Variable Labels / Value Labels
  // ------------------------------------------------------------------
  const INT64_TYPE = {
    RADIO: 'single',
    SELECT: 'single',
    RADIOSET: 'scale',
    RADIOSETS: '', // 같은 보기의 문항 묶음 → 척도/단일 여부는 보기로 자동 판단
    CHECK: 'multi01',
    HCHECK: 'multi01',
    CHECKSETS: 'multi01',
    GRADE_CLICK: 'rank',
    GRADE: 'rank',
    RANK: 'rank',
    NUMBER: 'numeric',
    HNUMBER: 'numeric',
    NUMBERSETS: 'numeric',
    TEXT: 'exclude',
    TEXTAREA: 'exclude',
    ADDR: 'exclude',
    EMAIL: 'exclude',
    PHONE: 'exclude',
  };

  function int64Detect(sheets) {
    let score = 0;
    if (findSheet(sheets, /^guide2$/i)) score += 5;
    if (findSheet(sheets, /^value\s*labels$/i)) score += 3;
    if (findSheet(sheets, /^variable\s*labels$/i)) score += 3;
    const g = findSheet(sheets, /^guide$/i);
    if (g && (sheets[g] || []).some((r) => r && /^(RADIO|CHECK|HCHECK|GRADE_CLICK|RADIOSETS)$/.test(str(r[3])))) score += 5;
    return score >= 5 ? 10 : 0;
  }

  function int64Parse(sheets) {
    const vars = new Map();
    const order = [];
    const warnings = [];
    const get = (name) => {
      const key = keyOf(name);
      if (!vars.has(key)) {
        vars.set(key, newVar(str(name)));
        order.push(key);
      }
      return vars.get(key);
    };
    const answers = new Map(); // GUIDE2의 변수별 보기

    // --- GUIDE2: ORD | COLUMNNAME | QUESTION | SUB_QUESTION | ANSWER | VALUE
    const g2name = findSheet(sheets, /^guide2$/i);
    if (g2name) {
      const rows = sheets[g2name];
      const h = rows[0] || [];
      const ci = {
        name: headerIndex(h, /^column\s*name$|^columnname$/i),
        q: headerIndex(h, /^question$/i),
        sub: headerIndex(h, /^sub_?question$/i),
        ans: headerIndex(h, /^answer$/i),
        val: headerIndex(h, /^value$/i),
      };
      if (ci.name < 0) ci.name = 1;
      rows.slice(1).forEach((r) => {
        const name = str(r[ci.name]);
        if (!name) return;
        const v = get(name);
        const q = str(r[ci.q]);
        const sub = ci.sub >= 0 ? str(r[ci.sub]) : '';
        const ans = ci.ans >= 0 ? toNum(r[ci.ans]) : null;
        const val = ci.val >= 0 ? str(r[ci.val]) : '';
        if (!v.question) {
          let text = q;
          if (sub && !/^\d+$/.test(sub)) text += ' ▶' + sub;
          if (ans == null && val) text += ' ▶' + val;
          v.question = text;
        }
        if (ans != null && val) {
          if (!answers.has(v.name)) answers.set(v.name, []);
          const list = answers.get(v.name);
          if (!list.some((c) => c.code === ans)) list.push({ code: ans, label: val });
        }
      });
    }

    // --- Value Labels: "/A2_1_1 A2_1_2" 다음 줄부터 "1 '특허권'"
    const vlName = findSheet(sheets, /^value\s*labels$/i);
    if (vlName) {
      let current = [];
      sheets[vlName].forEach((r) => {
        const cells = (r || []).map(str);
        const head = cells.find((c) => c.startsWith('/'));
        if (head) {
          current = head.slice(1).trim().split(/\s+/).filter(Boolean);
          current.forEach((n) => get(n));
          return;
        }
        for (const c of cells) {
          const m = /^(-?\d+(?:\.\d+)?)\s+['"](.*)['"]\s*\.?$/.exec(c);
          if (!m) continue;
          const code = Number(m[1]);
          current.forEach((n) => {
            const v = get(n);
            if (!v.codes.some((x) => x.code === code)) v.codes.push({ code, label: m[2].trim() });
          });
          break;
        }
      });
    }

    // --- Variable Labels: 변수명 | ' 문항 - 항목'
    const vbName = findSheet(sheets, /^variable\s*labels$/i);
    if (vbName) {
      sheets[vbName].slice(1).forEach((r) => {
        const name = str(r && r[0]);
        const label = str(r && r[1]).replace(/^['"]\s*|\s*['"]$/g, '').trim();
        if (!name || !label || name.startsWith('.')) return;
        const v = get(name);
        if (!v.question) v.question = label;
      });
    }

    // --- GUIDE: 변수명 | 문항 | 문항유형번호 | 문항유형  (+ 보기 행: 코드 | 레이블)
    const gName = findSheet(sheets, /^guide$/i);
    if (gName) {
      const inRange = new Set();
      let target = [];
      const expand = (spec) => {
        const m = /^(\S+)\s+TO\s+(\S+)$/i.exec(spec);
        if (!m) return [spec];
        const a = order.indexOf(keyOf(m[1]));
        const b = order.indexOf(keyOf(m[2]));
        if (a < 0 || b < a) return [m[1], m[2]];
        return order.slice(a, b + 1).filter((k) => !/_etc|#/i.test(k)).map((k) => vars.get(k).name);
      };
      sheets[gName].slice(1).forEach((r) => {
        if (!r) return;
        const c0 = r[0];
        if (typeof c0 === 'number' || (str(c0) !== '' && toNum(c0) != null && /^\s*-?\d+\s*$/.test(str(c0)))) {
          // 보기 행
          const code = toNum(c0);
          const label = str(r[1]);
          target.forEach((n) => {
            const v = get(n);
            if (!v.guideCodes) v.guideCodes = [];
            if (!v.guideCodes.some((x) => x.code === code)) v.guideCodes.push({ code, label });
          });
          return;
        }
        const spec = str(c0);
        if (!spec) return;
        const rawType = str(r[3]).toUpperCase();
        const type = rawType in INT64_TYPE ? INT64_TYPE[rawType] : '';
        const q = str(r[1]);
        const isRange = /\sTO\s/i.test(spec);
        target = expand(spec);
        if (isRange) {
          const gid = spec.toUpperCase();
          target.forEach((n) => {
            const v = get(n);
            inRange.add(keyOf(n));
            v.group = gid;
            if (q) v.groupTitle = q;
            if (rawType) {
              v.type = type;
              v.rawType = rawType;
            }
          });
        } else {
          const v = get(spec);
          if (!v.question && q) v.question = q;
          // 범위(TO)로 이미 정한 변수는 개별 행이 덮어쓰지 않음
          if (inRange.has(keyOf(spec)) || !rawType) return;
          if (!v.rawType || v.rawType === 'TEXT' || rawType !== 'TEXT') {
            v.type = type;
            v.rawType = rawType;
          }
        }
      });
    }

    // --- 마무리: 보기와 항목 이름 정리
    vars.forEach((v) => {
      const ans = answers.get(v.name) || [];
      if (!v.codes.length && ans.length > 1) v.codes = ans.slice();
      if (!v.codes.length && v.guideCodes && v.guideCodes.length) v.codes = v.guideCodes.slice();
      if (v.type === 'multi01') {
        // 복수응답 변수는 GUIDE2에 자기 보기 하나만 있음 → 항목 이름
        if (ans.length === 1) v.itemLabel = ans[0].label;
        else {
          const m = /_(\d+)$/.exec(v.name);
          const hit = m && v.codes.find((c) => c.code === Number(m[1]));
          if (hit) v.itemLabel = hit.label;
        }
      }
      delete v.guideCodes;
    });
    if (!vars.size) warnings.push('INT64 코드북에서 변수를 찾지 못했습니다.');
    return { vars, warnings, format: 'int64' };
  }

  registerFormat({
    id: 'int64',
    name: 'INT64',
    usesSheet: false,
    detect: int64Detect,
    parse: int64Parse,
  });

  // ------------------------------------------------------------------
  // 형식 자동 감지 + 읽기
  //   sheets: { 시트이름: 2차원 배열 }
  //   opts: { format: 'auto' | 형식 id, sheet: 표준 양식에서 쓸 시트 }
  // ------------------------------------------------------------------
  function detectFormat(sheets) {
    let best = FORMATS.find((f) => f.id === 'standard');
    let score = 1;
    FORMATS.forEach((f) => {
      const s = f.detect(sheets);
      if (s > score) {
        best = f;
        score = s;
      }
    });
    return best;
  }

  function parseCodebookBook(sheets, opts) {
    opts = opts || {};
    const fmt = !opts.format || opts.format === 'auto' ? detectFormat(sheets) : FORMATS.find((f) => f.id === opts.format) || detectFormat(sheets);
    const res = fmt.parse(sheets, opts);
    res.format = fmt.id;
    res.formatName = fmt.name;
    return res;
  }

  Object.assign(TG, { FORMATS, registerFormat, detectFormat, parseCodebookBook });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

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

  // HTML 흔적(&#39;, <br>)과 줄바꿈 정리
  function cleanText(s) {
    return str(s)
      .replace(/<br\s*\/?>/gi, ' ')
      .replace(/<[^>]+>/g, '')
      .replace(/&#39;|&apos;/g, "'")
      .replace(/&quot;/g, '"')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&nbsp;/g, ' ')
      .replace(/&amp;/g, '&')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // SPSS 문법 형태의 Value Labels 시트: "/변수1 변수2" 다음 줄부터 "1 '보기'" 또는 "1.'보기'"
  function readSpssValueLabels(rows, get) {
    let current = [];
    (rows || []).forEach((r) => {
      const cells = (r || []).map(str);
      const head = cells.find((c) => c.startsWith('/'));
      if (head) {
        current = Array.from(new Set(head.slice(1).trim().split(/\s+/).filter(Boolean)));
        current.forEach((n) => get(n));
        return;
      }
      for (const c of cells) {
        const m = /^(-?\d+(?:\.\d+)?)\s*[.:]?\s*['"]([\s\S]*)['"]\s*\.?$/.exec(c);
        if (!m) continue;
        const code = Number(m[1]);
        const label = cleanText(m[2]);
        current.forEach((n) => {
          const v = get(n);
          if (!v.codes.some((x) => x.code === code)) v.codes.push({ code, label });
        });
        break;
      }
    });
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
    if (vlName) readSpssValueLabels(sheets[vlName], get);

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
  // 3) G-CAII (사내 조사 시스템)
  //    시트: 질문지(문항ID | 문항타입 | 보기 | 로직 | page) / Variable Labels / Value Labels
  //    변수명 규칙(질문번호 뒤에 붙는 기호)
  //      M숫자 = 복수응답, MT숫자 = 행렬 문항의 행, MT숫자M숫자 = 행렬 복수응답,
  //      C숫자MT숫자 = 행렬 멀티, R숫자 = 순위, N숫자 = 숫자, O/P/E = 주관식·전화·메일,
  //      9997 = 기타, 9999 = 모름, 질문번호의 '-'는 'K' (B5-1 → B5K1)
  // ------------------------------------------------------------------
  function gcaiiQuestionSheet(sheets) {
    return Object.keys(sheets).find((n) => {
      const h = (sheets[n] || [])[0] || [];
      return headerIndex(h, /^문항\s*ID$/i) >= 0 && headerIndex(h, /^문항\s*타입$/) >= 0;
    });
  }

  function gcaiiDetect(sheets) {
    let score = 0;
    if (gcaiiQuestionSheet(sheets)) score += 6;
    if (findSheet(sheets, /^value\s*labels$/i)) score += 2;
    if (findSheet(sheets, /^variable\s*labels$/i)) score += 2;
    return score >= 6 ? 10 : 0;
  }

  const GCAII_QTYPE = {
    SIG: '', // 단수응답: 보기로 단일/척도 판단
    MTP: 'multi01',
    MX1: '',
    MX3: 'multi01',
    MX7: '',
    RNK: 'rank',
    NUM: 'numeric',
    NUD: 'numeric',
    OPN: 'exclude',
    TEL: 'exclude',
    EML: 'exclude',
    MSG: 'exclude',
  };

  // "단수응답형 (SIG)" 같은 약어와 "단순응답형(Single Answer)" 같은 전체 이름 모두 읽기
  function gcaiiQtype(raw) {
    const s = str(raw);
    const abbr = /\(([A-Z0-9]{2,4})\)\s*$/.exec(s);
    if (abbr && abbr[1] in GCAII_QTYPE) return abbr[1];
    if (/메시지|message/i.test(s)) return 'MSG';
    if (/순위|rank/i.test(s)) return 'RNK';
    if (/멀티|matrix\s*-?\s*multi\b.*level/i.test(s)) return 'MX7';
    if (/척도형?-?\s*복수|matrix\s*-\s*multi/i.test(s)) return 'MX3';
    if (/척도|matrix/i.test(s)) return 'MX1';
    if (/복수|multi/i.test(s)) return 'MTP';
    if (/전화|phone/i.test(s)) return 'TEL'; // "Phone Number"가 숫자형으로 잡히지 않도록 먼저 확인
    if (/이메일|e-?mail/i.test(s)) return 'EML';
    if (/쌍숫자/.test(s)) return 'NUD';
    if (/숫자|numeric|number/i.test(s)) return 'NUM';
    if (/자기기입|주관|open|text/i.test(s)) return 'OPN';
    if (/단수|단순|단일|single/i.test(s)) return 'SIG';
    return '';
  }

  function gcaiiParse(sheets) {
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

    // --- 질문지: 질문번호("B5-1]")와 문항타입, 행렬 문항의 행 이름("1]교통카드 …")
    const questions = [];
    const qName = gcaiiQuestionSheet(sheets);
    if (qName) {
      const rows = sheets[qName];
      const h = rows[0];
      const ci = { id: headerIndex(h, /^문항\s*ID$/i), type: headerIndex(h, /^문항\s*타입$/), text: headerIndex(h, /^보기$/) };
      if (ci.text < 0) ci.text = 2;
      let cur = null;
      let pending = null; // "1. " 처럼 이름이 비어 있는 보기 → 다음 줄의 글자가 보기 이름
      rows.slice(1).forEach((r) => {
        if (!r) return;
        const text = cleanText(r[ci.text]);
        if (str(r[ci.id])) {
          // 질문번호: "SQ1]", "B5-1]", "1]", "12-1]", "개인정보동의]"
          const m = /^\s*([^\]\s]{1,20})\]\s*([\s\S]*)$/.exec(text);
          cur = null;
          pending = null;
          if (!m) return;
          cur = { code: m[1], base: m[1].replace(/-/g, 'K').toUpperCase(), text: m[2].trim(), qtype: gcaiiQtype(r[ci.type]), rows: [], codes: [] };
          questions.push(cur);
          return;
        }
        if (!cur || !text) return;
        let m = /^(\d+)\](.+)$/.exec(text);
        if (m) {
          cur.rows.push({ n: Number(m[1]), label: m[2].trim() });
          return;
        }
        m = /^(-?\d+)\.(.*)$/.exec(text);
        if (m) {
          const c = { code: Number(m[1]), label: m[2].trim() };
          cur.codes.push(c);
          pending = c.label ? null : c;
          return;
        }
        if (pending && !/^-?\d+$/.test(text)) {
          pending.label = text;
          pending = null;
        }
      });
    }
    // 긴 질문번호부터 맞춰 봄 (B5K1 이 B5 보다 먼저)
    const byLen = questions.slice().sort((a, b) => b.base.length - a.base.length);
    // 질문번호에 쓰인 영문 머리(SQ, A, B …) — 이것과 다른 머리는 데이터 쪽 접두어(Y1, Q1 …)로 보고 떼어 냄
    const heads = new Set(questions.map((q) => (/^([A-Z]+)\d/.exec(q.base) || [])[1]).filter(Boolean));
    const fits = (up, base) => up.startsWith(base) && (up.length === base.length || !/\d/.test(up[base.length]));
    const matchQuestion = (name) => {
      const up = str(name).toUpperCase();
      let q = byLen.find((x) => fits(up, x.base));
      if (q) return { q, rest: up.slice(q.base.length) };
      const pm = /^([A-Z]{1,3})(\d.*)$/.exec(up);
      if (pm && !heads.has(pm[1])) {
        q = byLen.find((x) => /^\d/.test(x.base) && fits(pm[2], x.base));
        if (q) return { q, rest: pm[2].slice(q.base.length) };
      }
      return null;
    };
    const questionOf = (name) => {
      const hit = matchQuestion(name);
      return hit ? hit.q : null;
    };

    // --- Variable Labels: 변수명 | 문항 - 항목 | 단위/문구
    const vbName = findSheet(sheets, /^variable\s*labels$/i);
    const units = new Map();
    if (vbName) {
      sheets[vbName].slice(1).forEach((r) => {
        const name = str(r && r[0]);
        if (!name || name.startsWith('.')) return;
        const v = get(name);
        const label = cleanText(r[1]);
        if (label && !v.question) v.question = label.replace(/\s*-\s*_+\s*$/, '').replace(/\s*-\s*$/, '');
        const unit = cleanText(str(r[2]).split(/<br/i)[0]);
        if (unit && unit.length <= 10 && !/^[)\s]+$/.test(unit) && !/^※/.test(unit)) units.set(v.name, unit);
      });
    }

    // --- Value Labels
    const vlName = findSheet(sheets, /^value\s*labels$/i);
    if (vlName) readSpssValueLabels(sheets[vlName], get);

    // --- 변수명 규칙으로 유형·묶음 정하기
    const mx7cols = new Map(); // 행렬 멀티 문항의 열 개수
    // 밑줄형 변수명(SQ2_1, Y4_3, Y20_1순위)을 기호형(N1, MT3, R1)으로 맞춤
    const normalizeRest = (rest, qtype) => {
      if (/^_?9997(ET|_?TXT)?$/i.test(rest) || /_?\d*ET$/i.test(rest)) return '_TXT';
      let m = /^_(\d+)\s*순위$/.exec(rest);
      if (m) return 'R' + m[1];
      m = /^_(\d+)$/.exec(rest);
      if (m) {
        const tag = { MX1: 'MT', MTP: 'M', MX3: 'M', NUM: 'N', NUD: 'N', RNK: 'R' }[qtype];
        if (tag) return tag + m[1];
      }
      return rest;
    };
    const classify = (v, q, rest) => {
      const up = v.name.toUpperCase();
      const qtype = q.qtype;
      const prefix = `${q.code}. `;
      let m;
      rest = normalizeRest(rest, qtype);
      if (/_TXT$/i.test(rest) || /_TXT$/i.test(up) || /9997$/.test(rest) && /^(R|N|O)/.test(rest) || /N9999$/.test(rest)) {
        v.type = 'exclude';
      } else if (/^[OPE]\d+$/.test(rest)) {
        v.type = 'exclude';
      } else if ((m = /^MT(\d+)M(\d+)$/.exec(rest))) {
        // 행렬 복수응답: 같은 열(M숫자)끼리 묶어서 행(자료 이름)을 항목으로
        const col = Number(m[2]);
        const colLabel = (q.codes.find((c) => c.code === col) || {}).label || `응답 ${col}`;
        const row = q.rows.find((x) => x.n === Number(m[1]));
        v.type = 'multi01';
        v.group = `${q.base}MT*M${col}`;
        v.groupTitle = `${prefix}${q.text} - ${colLabel}`;
        if (row) v.itemLabel = row.label;
        v.baseAll = true;
      } else if ((m = /^C(\d+)MT(\d+)$/.exec(rest))) {
        // 행렬 멀티: 같은 열(MT숫자)끼리 묶어서 행(C숫자)을 항목으로
        const col = Number(m[2]);
        const row = q.rows.find((x) => x.n === Number(m[1]));
        // 열 개수는 Variable Labels에 실제로 있는 변수로만 셈 (Value Labels에는 없는 변수도 섞여 있음)
        if (v.question) mx7cols.set(q.base, Math.max(mx7cols.get(q.base) || 0, col));
        v.type = '';
        v.group = `${q.base}C*MT${col}`;
        v.mx7 = { q, col };
        if (row) v.itemLabel = row.label;
      } else if ((m = /^MT(\d+)$/.exec(rest))) {
        const row = q.rows.find((x) => x.n === Number(m[1]));
        v.type = '';
        v.group = `${q.base}MT`;
        v.groupTitle = prefix + q.text;
        if (row) v.itemLabel = row.label;
      } else if (/^M\d+$/.test(rest)) {
        v.type = 'multi01';
        v.group = `${q.base}M`;
        v.groupTitle = prefix + q.text;
        const hit = /(\d+)$/.exec(rest) && v.codes.find((c) => c.code === Number(/(\d+)$/.exec(rest)[1]));
        if (hit) v.itemLabel = hit.label;
      } else if (/^R\d+$/.test(rest)) {
        v.type = 'rank';
        v.group = `${q.base}R`;
        v.groupTitle = prefix + q.text;
      } else if (/^N\d+$/.test(rest)) {
        v.type = 'numeric';
        v.codes = [];
      } else {
        v.type = GCAII_QTYPE[qtype] != null ? GCAII_QTYPE[qtype] : '';
        if (v.type === 'numeric') v.codes = [];
      }
      if (v.type === 'numeric' && units.has(v.name)) v.question += ` (단위: ${units.get(v.name)})`;
    };
    vars.forEach((v) => {
      const hit = matchQuestion(v.name);
      if (!hit) return;
      if (v.question) v.question = `${hit.q.code}. ${v.question}`;
      classify(v, hit.q, hit.rest);
    });

    // Variable Labels가 없거나 빠진 변수: 데이터 변수명을 질문번호에 맞춰 질문지만으로 만듦
    const resolve = (name) => {
      const hit = matchQuestion(name);
      if (!hit || hit.q.qtype === 'MSG') return null;
      const { q, rest } = hit;
      const v = newVar(str(name));
      v.codes = q.codes.filter((c) => c.label && !/^_+$/.test(c.label)).map((c) => ({ code: c.code, label: c.label }));
      v.question = `${q.code}. ${q.text}`;
      // 숫자 여러 칸(년/개월 등)은 칸 이름을 제목에 붙임
      const nm = /^_?N?(\d+)$/.exec(normalizeRest(rest, q.qtype).replace(/^N/, '_'));
      if (/^(NUM|NUD)$/.test(q.qtype) && nm) {
        const part = q.codes.find((c) => c.code === Number(nm[1]));
        const label = part ? part.label.replace(/_+/g, '').trim() : '';
        if (label) v.question += ` - ${label}`;
      }
      classify(v, q, rest);
      return v;
    };

    // 행렬 멀티(MX7): 보기(1 예, 2 아니오, 3 예, 4 아니오)를 열마다 나누고, 열 제목은 질문 문장에서
    vars.forEach((v) => {
      if (!v.mx7) return;
      const { q, col } = v.mx7;
      const ncol = mx7cols.get(q.base) || 1;
      const all = (q.codes.length ? q.codes : v.codes).slice().sort((a, b) => a.code - b.code);
      const per = Math.floor(all.length / ncol);
      if (per >= 2 && per * ncol === all.length) v.codes = all.slice((col - 1) * per, col * per);
      const sentences = q.text.split(/(?<=[?？])\s*/).map((x) => x.trim()).filter(Boolean);
      const colTitle = sentences.length === ncol ? sentences[col - 1] : `${q.text} [응답 ${col}]`;
      v.groupTitle = `${q.code}. ${colTitle}`;
      delete v.mx7;
    });

    if (!questions.length) warnings.push('G-CAII 질문지 시트에서 문항을 찾지 못했습니다.');
    return { vars, warnings, format: 'gcaii', resolve };
  }

  registerFormat({
    id: 'gcaii',
    name: 'G-CAII',
    usesSheet: false,
    detect: gcaiiDetect,
    parse: gcaiiParse,
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

  Object.assign(TG, { FORMATS, registerFormat, detectFormat, parseCodebookBook, cleanText });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

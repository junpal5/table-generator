/*
 * 화면에서 만든 표와 같은 표를 SPSS에서 뽑을 수 있는 신택스(.sps)를 만듭니다.
 * 사내 TABLES 신택스 형식을 따릅니다.
 *   머리 주석 → COMPUTE X=1 → WEIGHT → 배너(T_/V_) 정의 → 파생 변수(평균·척도·요약)
 *   → VARIABLE/VALUE LABELS → 응답자 특성 표 → 표마다 TABLE 명령 → EXECUTE
 * 파일 경로(CD, GET FILE)와 저장(SAVE OUTFILE)은 넣지 않습니다.
 */
(function (root) {
  'use strict';
  const TG = (root.TG = root.TG || {});

  const PTO = "/PTO=T1'■ 전      체 ■' T2''/FTO=T3'계'";
  const FORMAT = "/FORMAT ZERO MISSING('.')";

  // SPSS 문자열: 작은따옴표 안의 작은따옴표는 두 번
  const q = (s) => "'" + String(s == null ? '' : s).replace(/\s+/g, ' ').replace(/'/g, "''") + "'";
  const dq = (s) => '"' + String(s == null ? '' : s).replace(/\s+/g, ' ').replace(/"/g, '""') + '"';
  const clip = (s, n) => {
    s = String(s == null ? '' : s).replace(/\s+/g, ' ').trim();
    return s.length > n ? s.slice(0, n - 1) + '…' : s;
  };
  // SPSS 변수 이름으로 쓸 수 있는지 (글자·숫자·_ . @ # $, 첫 글자는 글자/@/#/$)
  const validName = (v) => /^[A-Za-z가-힣@#$][A-Za-z0-9가-힣_.@#$]*$/.test(v) && !/\.$/.test(v) && v.length <= 64;
  // 파생 변수 이름 (T_Q1, AA_Q1 …)
  const dn = (prefix, v) => prefix + String(v).replace(/[^A-Za-z0-9가-힣_@#$]/g, '_');

  function codeNum(c) {
    return Number.isInteger(c) ? String(c) : String(c);
  }

  // ------------------------------------------------------------------
  function buildSyntax(result, items, data, opts) {
    const out = [];
    const L = (s) => out.push(s == null ? '' : s);
    const itemById = new Map(items.map((it) => [it.id, it]));
    const labelled = new Set(); // VALUE LABELS를 이미 쓴 변수
    const warnings = [];
    const byItem = new Map(); // itemId → 표 목록(순서대로)
    result.tables.forEach((t) => {
      if (!byItem.has(t.itemId)) byItem.set(t.itemId, []);
      byItem.get(t.itemId).push(t);
    });
    const used = items.filter((it) => byItem.has(it.id));

    // ---------------- 머리
    const now = new Date();
    L('**.');
    L('*   코    드   :');
    L('*   프로젝트   :');
    L('*   담당부서   :');
    L(`*   작 성 일   :     ${now.getFullYear()}년 ${now.getMonth() + 1}월`);
    L('*   연 구 원   :');
    L('*   D      P   :');
    L('*   설문 통계표 생성기에서 만든 신택스입니다. 데이터 파일을 먼저 연 뒤 실행하세요.');
    L('**.');
    L('COMPUTE X=1.');
    L('**.');
    L('');
    L(opts.weightVar ? `WEIGHT BY ${opts.weightVar}.` : 'WEIGHT OFF.');
    L('');

    // ---------------- 배너(DEMO) 정의
    const banners = (opts.banners || []).map((b) => {
      const col = data.columns.get(TG.keyOf(b.var));
      return { ...b, name: col ? col.name : b.var, isText: !!(col && col.raw) };
    });
    const T = banners.map((b) => dn('T_', b.name));
    const V = banners.map((b) => dn('V_', b.name));
    const tabRows = ['T1'].concat(T).join('+');
    const cpcList = T.join(' ');
    L('**.');
    L('*기본 DEMO 정의.');
    L('**.');
    banners.forEach((b, i) => {
      if (b.isText) {
        // 문자로 된 배너(예: 지역 "서울")는 숫자 코드로 바꿔서 사용
        const pairs = b.codes.map((c) => `(${q(c.label)}=${codeNum(c.code)})`).join('');
        L(`RECODE ${b.name} ${pairs} INTO ${T[i]}.`);
        L(`RECODE ${b.name} ${pairs} INTO ${V[i]}.`);
      } else {
        if (!validName(b.name)) warnings.push(`배너 변수명 '${b.name}'은 SPSS 변수명 규칙에 맞지 않을 수 있습니다.`);
        L(`COMPUTE ${T[i]}=${b.name}.`);
        L(`COMPUTE ${V[i]}=${b.name}.`);
      }
    });
    banners.forEach((b, i) => {
      L(`VARIABLE LABEL ${T[i]} ${q(b.label)}.`);
      L(`VARIABLE LABEL ${V[i]} ${q(b.label)}.`);
    });
    L('');
    banners.forEach((b, i) => {
      if (b.isText) return;
      L(`RECODE ${T[i]} ${V[i]}`);
      b.codes.forEach((c) => L(`(${codeNum(c.code)}=${codeNum(c.code)})`));
      L('(ELSE=SYSMIS).');
      L('');
    });
    banners.forEach((b, i) => {
      L(`VALUE LABEL ${T[i]} ${V[i]}`);
      b.codes.forEach((c) => L(`${codeNum(c.code)} ${q(clip(c.label, 60))}`));
      L('.');
      L('');
    });

    // ---------------- 파생 변수 (평균·척도·요약)
    const derived = [];
    const D = (s) => derived.push(s);
    const scaleOf = new Map();

    used.forEach((it) => {
      const tables = byItem.get(it.id);
      if (it.kind === 'numeric') {
        const v = it.vars[0];
        const aa = dn('AA_', v);
        const c = dn('C_', v);
        D(`COMPUTE ${aa}=${v}.`);
        // 9999·999999 같은 모름 코드는 평균에서 제외 (화면과 같은 기준)
        const nines = Array.from(new Set(data.columns.get(it.keys[0]).values.filter((x) => x != null && TG.NINES.test(String(x)))));
        nines.forEach((n) => D(`IF (${v} EQ ${n}) ${aa}=$SYSMIS.`));
        D(`COMPUTE ${c}=${aa}.`);
        D(`RECODE ${c}(LO THRU HI=1).`);
        D(`VARIABLE LABEL ${aa} '(평균)'.`);
        D(`VARIABLE LABEL ${c} '사례수'.`);
        D('');
      }
      if (it.kind === 'scale' || it.kind === 'scaleset') {
        it.vars.forEach((v) => {
          const spec = TG.scaleSpec(it.codes);
          if (!spec) return;
          scaleOf.set(v, spec);
          const aa = dn('AA_', v);
          const y = dn('Y_', v);
          const g = dn('G_', v);
          const rng = spec.max - spec.min;
          D(`*${v}: ${spec.points}점 척도 평균${spec.positiveHigh ? '' : '(긍정일수록 높은 점수가 되도록 역산)'}.`);
          const cond = `(${v} GE ${spec.min} AND ${v} LE ${spec.max})`;
          D(spec.positiveHigh ? `IF ${cond} ${aa}=${v}.` : `IF ${cond} ${aa}=${spec.min + spec.max}-${v}.`);
          D(`IF ${cond} ${y}=(${aa}-${spec.min})*100/${rng}.`);
          D(`VARIABLE LABEL ${aa} '(평균)'.`);
          D(`VARIABLE LABEL ${y} '100점 평균'.`);
          // 긍정·보통·부정 묶음
          const top = Array.from(spec.topSet).join(' ');
          const bot = Array.from(spec.botSet).join(' ');
          const mid = Array.from(spec.validSet).filter((c) => !spec.topSet.has(c) && !spec.botSet.has(c));
          D(`COMPUTE ${g}=${v}.`);
          D(`RECODE ${g}(${bot}=21)${mid.length ? `(${mid.join(' ')}=22)` : ''}(${top}=23)(ELSE=SYSMIS).`);
          D(`VALUE LABEL ${g}`);
          D(`21 '⊙ ${spec.botName}'`);
          if (mid.length) D(`22 '⊙ 보통${spec.points === 11 ? `(${mid[0]}~${mid[mid.length - 1]}점)` : ''}'`);
          D(`23 '⊙ ${spec.topName}'`);
          D('.');
          D('');
        });
      }
      // 요약표용 변수: 첫 항목 사례수(C_), 보기 비율(S1_, S2_ …)
      const summaries = tables.filter((t) => t.kind === 'summary');
      if (summaries.length) {
        const c = dn('C_', it.vars[0]);
        D(`*요약표용 변수(${it.vars[0]} ~ ${it.vars[it.vars.length - 1]}).`);
        D(`COMPUTE ${c}=${it.vars[0]}.`);
        D(`RECODE ${c}(LO THRU HI=1).`);
        D(`VARIABLE LABEL ${c} '사례수'.`);
        if (it.kind === 'singleset') {
          it.codes.slice(0, -1).forEach((code, ci) => {
            it.vars.forEach((v, j) => {
              const s = dn(`S${ci + 1}_`, v);
              D(`COMPUTE ${s}=${v}.`);
              D(`RECODE ${s}(${codeNum(code.code)}=100)(MIS=SYSMIS)(ELSE=0).`);
              D(`VARIABLE LABEL ${s} ${q(clip(it.itemLabels[j], 120))}.`);
            });
          });
        } else if (it.kind === 'scaleset') {
          it.vars.forEach((v, j) => {
            const spec = scaleOf.get(v);
            if (!spec) return;
            const s = dn('T2_', v);
            D(`IF (${v} GE ${spec.min} AND ${v} LE ${spec.max}) ${s}=0.`);
            D(`IF (ANY(${v},${Array.from(spec.topSet).join(',')})) ${s}=100.`);
            D(`VARIABLE LABEL ${s} ${q(clip(it.itemLabels[j], 120))}.`);
            D(`VARIABLE LABEL ${dn('AA_', v)} ${q(clip(it.itemLabels[j], 120))}.`);
            D(`VARIABLE LABEL ${dn('Y_', v)} ${q(clip(it.itemLabels[j], 120))}.`);
          });
        }
        D('');
      }
    });
    if (derived.length) {
      L('**.');
      L('*평균·척도·요약 변수 --> 무응답은 SYSMIS로 처리.');
      L('**.');
      derived.forEach(L);
    }

    // ---------------- 문항 변수·값 레이블
    const varLabels = [];
    const valueLabels = [];
    used.forEach((it) => {
      if (it.kind === 'multi01') {
        const positional = it.keys.some((k) => data.columns.get(k).values.some((x) => x != null && x !== 0 && x !== 1));
        const pairs = [];
        it.vars.forEach((v, j) => {
          varLabels.push(`VARIABLE LABELS ${v} ${q(clip(it.itemLabels[j], 120))}.`);
          // 선택하면 보기 번호가 들어가는 방식: 값 = 보기 번호
          const val = positional ? data.columns.get(it.keys[j]).values.find((x) => x != null && x !== 0) : null;
          if (val != null) pairs.push([val, it.itemLabels[j]]);
        });
        // TABLES의 /MRG는 묶음의 첫 변수 값 레이블을 쓰므로, 묶음 전체에 같은 레이블 목록을 붙임
        if (pairs.length) {
          valueLabels.push('/' + it.vars.join(' '));
          pairs.sort((a, b) => a[0] - b[0]).forEach(([val, lab]) => valueLabels.push(`\t${codeNum(val)} ${q(clip(lab, 60))}`));
          it.vars.forEach((v) => labelled.add(v));
        }
        return;
      }
      if (it.kind === 'numeric' || !it.codes.length) {
        it.vars.forEach((v) => varLabels.push(`VARIABLE LABELS ${v} ${q(clip(it.title, 120))}.`));
        return;
      }
      it.vars.forEach((v, j) => {
        const title = it.isGroup && it.kind !== 'rank' && it.kind !== 'multicode' ? `${it.title} - ${it.itemLabels[j]}` : it.title;
        varLabels.push(`VARIABLE LABELS ${v} ${q(clip(title, 120))}.`);
      });
      const todo = it.vars.filter((v) => !labelled.has(v));
      if (!todo.length) return;
      todo.forEach((v) => labelled.add(v));
      valueLabels.push('/' + todo.join(' '));
      it.codes.forEach((c) => valueLabels.push(`\t${codeNum(c.code)} ${q(clip(c.label, 60))}`));
    });
    L('**.');
    L('*문항 변수 레이블.');
    L('**.');
    varLabels.forEach(L);
    L('');
    if (valueLabels.length) {
      L('VALUE LABELS');
      valueLabels.forEach(L);
      L('.');
      L('');
    }
    used.forEach((it) => it.vars.forEach((v) => {
      if (!validName(v)) warnings.push(`문항 변수명 '${v}'은 SPSS 변수명 규칙에 맞지 않을 수 있습니다.`);
    }));

    // ---------------- 표
    const banner = (by, sta, extra, corner, title) => {
      L('TABLE');
      L(FORMAT);
      L(PTO);
      (extra || []).forEach(L);
      L(`/TAB=${tabRows} BY ${by}`);
      sta.forEach((s, i) => L((i === 0 ? '/STA=' : '     ') + s));
      L(`/CORNER=${dq(clip('BASE=' + corner, 100))}`);
      L('/TITLE=');
      L(dq(clip(title, 120)));
      L('');
    };
    const cpc = (v) => `CPC(${v}(F4.1)'%':${cpcList})`;
    const cou = "COU(T2(PAREN6.0)'사례수')";

    L('**.');
    L('*@@@@@ TABLE.');
    L('**.');
    if (banners.length) {
      L('TABLE');
      L(FORMAT);
      L(PTO);
      L(`/TAB=${tabRows} BY X`);
      L("/STA=COU((PAREN5.0)'사례수') CPC((F4.1)'%')");
      L('/TITLE=');
      L("'응 답 자  특 성'");
      L('');
    }

    const cornerOf = (t) => {
      const base = (t.rows.find((r) => r.isTotal) || t.rows[0] || {}).n;
      if (t.kind === 'multi' && itemById.get(t.itemId).baseAll) return '전체';
      return base === data.n ? '전체' : '해당 문항 응답자';
    };

    used.forEach((it) => {
      const tables = byItem.get(it.id);
      const vars = it.vars;
      tables.forEach((t, ti) => {
        L(`* NO=${t.no}.`);
        const title = `표 ${t.no}. ${t.title}`;
        const corner = cornerOf(t);
        if (t.kind === 'error') {
          L(`* 계산 오류로 제외: ${t.title}.`);
          L('');
          return;
        }
        if ((it.kind === 'single' || it.kind === 'singleset') && t.kind !== 'summary') {
          const v = it.kind === 'single' ? vars[0] : vars[ti];
          banner(`T2+${v}+T3`, [cou, cpc(v)], [], corner, title);
          return;
        }
        if (it.kind === 'scale' || (it.kind === 'scaleset' && t.kind !== 'summary')) {
          const v = it.kind === 'scale' ? vars[0] : vars[ti];
          const spec = scaleOf.get(v);
          if (!spec) {
            banner(`T2+${v}+T3`, [cou, cpc(v)], [], corner, title);
            return;
          }
          const aa = dn('AA_', v);
          const y = dn('Y_', v);
          banner(
            `T2+M+T3+${aa}+${y}`,
            [cou, cpc('M'), `MEA(${aa}(PAREN7.2)'점')`, `MEA(${y}(PAREN7.2)'점')`],
            [`/MRG=M''${v} ${dn('G_', v)}`, `/OBS=${aa} ${y}`],
            corner,
            title,
          );
          return;
        }
        if (it.kind === 'numeric') {
          const v = vars[0];
          const aa = dn('AA_', v);
          const c = dn('C_', v);
          banner(
            `${c}+${aa}`,
            [`SUM(${c}(PAREN6.0)'사례수')`, `MEA(${aa}(F8.2)'평균')`, `STDDEV(${aa}(F8.2)'표준편차')`, `MINIMUM(${aa}(F8.2)'최솟값')`, `MEDIAN(${aa}(F8.2)'중앙값')`, `MAXIMUM(${aa}(F8.2)'최댓값')`],
            [`/OBS=${c} ${aa}`],
            corner,
            title,
          );
          return;
        }
        if (it.kind === 'multi01') {
          const positional = it.keys.some((k) => data.columns.get(k).values.some((x) => x != null && x !== 0 && x !== 1));
          // 0/1 방식은 1을 센 복수 이분형 묶음(MDG), 보기 번호 방식은 복수 범주형 묶음(MRG)
          const grp = positional ? `/MRG=M''${vars.join(' ')}` : `/MDG=M''${vars.join(' ')}`;
          banner('T2+M', [cou, cpc('M')], [grp], corner, title);
          if (it.baseAll) L('* 주의: 화면에서는 전체 응답자를 Base로 계산했습니다. TABLES의 복수응답 %는 하나 이상 응답한 사람이 Base입니다.');
          return;
        }
        if (it.kind === 'rank') {
          if (ti === 0) banner(`T2+${vars[0]}+T3`, [cou, cpc(vars[0])], [], corner, title);
          else banner('T2+M', [cou, cpc('M')], [`/MRG=M''${vars.join(' ')}`], corner, title);
          return;
        }
        if (it.kind === 'multicode') {
          banner('T2+M', [cou, cpc('M')], [`/MRG=M''${vars.join(' ')}`], corner, title);
          return;
        }
        if (t.kind === 'summary') {
          const c = dn('C_', vars[0]);
          let sv;
          let fmt = "(F6.1)'%'";
          const sumIndex = tables.filter((x) => x.kind === 'summary').indexOf(t);
          if (it.kind === 'singleset') sv = vars.map((v) => dn(`S${sumIndex + 1}_`, v));
          else if (/100점/.test(t.title)) {
            sv = vars.map((v) => dn('Y_', v));
            fmt = "(F7.2)'점'";
          } else if (/평균/.test(t.title)) {
            sv = vars.map((v) => dn('AA_', v));
            fmt = "(F7.2)'점'";
          } else sv = vars.map((v) => dn('T2_', v));
          banner(
            [c].concat(sv).join('+'),
            [`SUM(${c}(PAREN6.0)'사례수')`].concat(sv.map((s) => `MEA(${s}${fmt})`)),
            [`/OBS=${sv.concat([c]).join(' ')}`],
            corner,
            title,
          );
          return;
        }
        L(`* 이 표 형식은 신택스로 만들지 못했습니다: ${t.title}.`);
        L('');
      });
    });
    if (opts.sortedAny) L('* 참고: 화면의 "큰 순 정렬"은 신택스에 반영되지 않습니다(보기 순서는 코드 순서).');
    L('');
    L('**.');
    L('EXECUTE.');
    L('**.');

    if (warnings.length) {
      out.splice(8, 0, ...Array.from(new Set(warnings)).map((w) => `*   확인 필요: ${w}`));
    }
    return out.join('\r\n') + '\r\n';
  }

  // EUC-KR(CP949) 또는 UTF-8로 저장
  function encodeSyntax(text, encoding) {
    const cp = root.cptable;
    if (encoding === 'euc-kr' && cp && cp.utils) {
      const body = '* Encoding: EUC-KR.\r\n' + text;
      return new Uint8Array(cp.utils.encode(949, body));
    }
    const body = '﻿* Encoding: UTF-8.\r\n' + text;
    return new TextEncoder().encode(body);
  }

  function downloadSyntax(result, items, data, opts, filename, encoding) {
    const bytes = encodeSyntax(buildSyntax(result, items, data, opts), encoding);
    const blob = new Blob([bytes], { type: 'text/plain' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename || '통계표.sps';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      URL.revokeObjectURL(a.href);
      a.remove();
    }, 1000);
  }

  Object.assign(TG, { buildSyntax, encodeSyntax, downloadSyntax });
  if (typeof module !== 'undefined' && module.exports) module.exports = TG;
})(typeof window !== 'undefined' ? window : globalThis);

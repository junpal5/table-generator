// 계산 엔진 검증 테스트. 실행: npm test
const assert = require('assert');
const path = require('path');
const XLSX = require('xlsx');
const TG = require('../js/core.js');
require('../js/render.js');

const SAMPLES = path.join(__dirname, '..', 'samples');
const rows = (f) => {
  const wb = XLSX.readFile(path.join(SAMPLES, f));
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
};

let passed = 0;
function test(name, fn) {
  try {
    fn();
    passed++;
    console.log('  ✓ ' + name);
  } catch (e) {
    console.error('  ✗ ' + name + '\n    ' + e.message);
    process.exitCode = 1;
  }
}
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 1e-9, `${msg}: ${a} != ${b}`);

const cb = TG.parseCodebook(rows('예시_코드북.xlsx'));
const data = TG.prepareData(rows('예시_데이터.xlsx'));
const { items } = TG.buildItems(cb, data);
const byVar = (v) => items.find((i) => i.vars[0] === v);
const col = (v) => data.columns.get(v.toLowerCase()).values;

console.log('자동 인식');
test('유형 자동 인식', () => {
  const kinds = Object.fromEntries(items.map((i) => [i.vars[0], i.kind]));
  assert.deepStrictEqual(kinds, {
    ID: 'exclude', SQ1: 'single', SQ2: 'single', SQ3: 'single', Q1: 'scale', Q2_1: 'scaleset',
    Q3_1: 'multi01', Q4_1: 'rank', Q5: 'numeric', Q6: 'single', Q7: 'exclude', wt: 'weight',
  });
});
test('묶음 제목과 항목 이름', () => {
  assert.strictEqual(byVar('Q2_1').title, 'Q2. 서비스 평가');
  assert.deepStrictEqual(byVar('Q4_1').itemLabels, ['1순위', '2순위']);
  assert.strictEqual(byVar('Q4_1').title, 'Q4. 서비스 선택 이유');
  assert.strictEqual(byVar('Q3_1').itemLabels[1], 'SNS');
});
test('배너 짧은 이름', () => {
  assert.strictEqual(TG.shortName('SQ1. 귀하의 성별은 무엇입니까?', 'SQ1'), '성별');
  assert.strictEqual(TG.shortName('SQ2. 귀하의 연령은 어떻게 되십니까?', 'SQ2'), '연령');
});

console.log('코드북 형식');
test('한 칸에 모은 보기(1=남자, 2=여자)', () => {
  const c = TG.parseCodebook([
    ['변수명', '변수 레이블', '값 레이블'],
    ['A1', '성별', '1=남자, 2=여자'],
    ['A2', '만족도', '1) 매우 불만족 2) 불만족 3) 보통 4) 만족 5) 매우 만족'],
  ]);
  assert.deepStrictEqual(c.vars.get('a1').codes, [{ code: 1, label: '남자' }, { code: 2, label: '여자' }]);
  assert.strictEqual(c.vars.get('a2').codes.length, 5);
  assert.strictEqual(c.vars.get('a2').codes[4].label, '매우 만족');
});
test('0/1 없이 1/빈칸으로 입력한 복수응답', () => {
  const d = TG.prepareData([['M_1', 'M_2', 'M_3'], [1, null, 1], [null, 1, null], [null, null, null]]);
  const r = TG.buildItems(null, d);
  assert.strictEqual(r.items[0].kind, 'multi01');
  const t = TG.computeTables(r.items, d, {}).tables[0];
  assert.strictEqual(t.rows[0].n, 2); // 하나도 선택 안 한 사람은 제외
  near(t.rows[0].values[0], 50, 'M_1 %');
});
test('같은 보기를 쓰는 개별 문항은 순위로 잡지 않음', () => {
  const d = TG.prepareData([['B_1', 'B_2', 'B_3'], [1, 1, 2], [2, 1, 1], [1, 2, 2], [2, 2, 1]]);
  const c = TG.parseCodebook([['변수명', '문항', '코드', '레이블'], ['B_1', '이용여부 - A', 1, '있다'], ['', '', 2, '없다'], ['B_2', '이용여부 - B', 1, '있다'], ['', '', 2, '없다'], ['B_3', '이용여부 - C', 1, '있다'], ['', '', 2, '없다']]);
  assert.strictEqual(TG.buildItems(c, d).items[0].kind, 'singleset');
});

console.log('표 계산 (가중치 없음)');
const plain = TG.computeTables(items, data, { banners: [{ var: 'SQ1', label: '성별' }] });
const tbl = (title) => plain.tables.find((t) => t.title === title);
test('단일응답 % 와 사례수', () => {
  const t = tbl('SQ1. 귀하의 성별은 무엇입니까?');
  const v = col('SQ1');
  const male = v.filter((x) => x === 1).length;
  assert.strictEqual(t.rows[0].n, 500);
  near(t.rows[0].values[0], (male / 500) * 100, '남자 %');
  assert.strictEqual(t.rows[1].label, '남자');
  assert.strictEqual(t.rows[1].n, male);
});
test('척도 평균·Top2 (모름 9는 평균에서 제외)', () => {
  const t = tbl('Q1. 전반적인 서비스 만족도');
  const v = col('Q1').filter((x) => x != null);
  const valid = v.filter((x) => x >= 1 && x <= 5);
  const mean = valid.reduce((a, b) => a + b, 0) / valid.length;
  const idx = (l) => t.columns.findIndex((c) => c.label === l);
  near(t.rows[0].values[idx('평균')], mean, '평균');
  near(t.rows[0].values[idx('100점 환산')], ((mean - 1) / 4) * 100, '100점');
  near(t.rows[0].values[idx('긍정(Top2)')], (v.filter((x) => x >= 4 && x <= 5).length / v.length) * 100, 'Top2');
});
test('순위 1+2순위 종합', () => {
  const t = tbl('Q4. 서비스 선택 이유 - 1+2순위');
  const a = col('Q4_1');
  const b = col('Q4_2');
  const base = a.filter((x) => x != null).length;
  const hits = a.filter((x, i) => x === 1 || b[i] === 1).length;
  near(t.rows[0].values[0], (hits / base) * 100, '가격 1+2순위');
});
test('스킵 문항은 응답자만 Base', () => {
  const t = tbl('Q6. (Q1=1,2 응답자) 불만족 이유');
  assert.strictEqual(t.rows[0].n, col('Q6').filter((x) => x != null).length);
  assert.ok(t.rows[0].n < 500);
});
test('수치 문항 평균', () => {
  const t = tbl('Q5. 최근 1개월 이용 횟수(회)');
  const v = col('Q5').filter((x) => x != null);
  near(t.rows[0].values[0], v.reduce((a, b) => a + b, 0) / v.length, '평균');
});

console.log('가중치');
test('가중 % = 가중치 합 비율', () => {
  const r = TG.computeTables(items.filter((i) => i.vars[0] === 'SQ1'), data, { weightVar: 'wt' });
  const w = col('wt');
  const s = col('SQ1');
  let m = 0;
  let tot = 0;
  s.forEach((x, i) => { tot += w[i]; if (x === 1) m += w[i]; });
  near(r.tables[0].rows[0].values[0], (m / tot) * 100, '가중 남자 %');
  near(r.tables[0].rows[0].wn, tot, '가중 사례수');
});

console.log('표 모양');
test('보고서형/배너형 격자', () => {
  const t = tbl('SQ1. 귀하의 성별은 무엇입니까?');
  const g1 = TG.toGrid(t, { orientation: 'row' });
  assert.strictEqual(g1.length, 1 + 3);
  const g2 = TG.toGrid(t, { orientation: 'col' });
  assert.strictEqual(g2[0][1].v, '전체');
  assert.strictEqual(g2.length, 2 + 1 + 3); // 머리글 2 + 사례수 + 보기 2 + 계
});

console.log(`\n${passed}개 통과`);

// 계산 엔진 검증 테스트. 실행: npm test
const assert = require('assert');
const path = require('path');
const XLSX = require('xlsx');
const TG = require('../js/core.js');
require('../js/formats.js');
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

console.log('INT64 형식');
// 실제 INT64 코드북과 같은 구조의 작은 가상 예시 (실제 조사 데이터는 저장소에 넣지 않음)
const int64Book = {
  GUIDE: [
    ['변수명', '문항', '문항유형번호', '문항유형'],
    ['A1_1', '기업명', 51, 'TEXT'],
    ['A1_3', '설립연도', 51, 'TEXT'],
    ['A1_3', null, null, 'NUMBER'],
    ['A1_6', '기업형태', 11, 'RADIO'],
    [1, '독립기업'], [2, '계열사'],
    ['A2_1_1 TO A2_1_3', '보유 산업재산권(모두 응답)', 21, 'CHECK'],
    [1, '특허권'], [2, '실용신안권'], [3, '디자인권'],
    ['A4_1A_1', '인지 여부▶사업A', 61, 'RADIOSETS'], [1, '알고 있다'], [2, '모른다'],
    ['A4_1A_2', '인지 여부▶사업B', 61, 'RADIOSETS'], [1, '알고 있다'], [2, '모른다'],
    ['A5_2_1', '[1순위] 필요한 정책', 32, 'GRADE_CLICK'],
    ['A5_2_1 TO A5_2_2', '[1-2순위] 필요한 정책', 32, 'GRADE_CLICK'],
    [1, '비용 지원'], [2, '인력 지원'], [3, '정보 제공'],
    ['A6_2', '보호 수준(0~10점)', 13, 'RADIOSET'],
    [0, '매우 약함'], [1, '<--'], [2, '---'], [3, '---'], [4, '---'], [5, '보 통'], [6, '---'], [7, '---'], [8, '---'], [9, '-->'], [10, '매우 강함'],
  ],
  GUIDE2: [
    ['ORD', 'COLUMNNAME', 'QUESTION', 'SUB_QUESTION', 'ANSWER', 'VALUE'],
    ['1', 'A1_1', '기업명', '', '', ''],
    ['2', 'A1_3', '설립연도', '', '', ''],
    ['3', 'A1_6', '기업형태', '', '1', '독립기업'],
    ['4', 'A1_6', '기업형태', '', '2', '계열사'],
    ['5', 'A2_1_1', '보유 산업재산권(모두 응답)', '', '1', '특허권'],
    ['6', 'A2_1_2', '보유 산업재산권(모두 응답)', '', '2', '실용신안권'],
    ['7', 'A2_1_3', '보유 산업재산권(모두 응답)', '', '3', '디자인권'],
    ['8', 'A4_1A_1', '인지 여부', '사업A', '1', '알고 있다'],
    ['9', 'A4_1A_2', '인지 여부', '사업B', '1', '알고 있다'],
    ['10', 'A5_2_1', '필요한 정책', '1', '1', '비용 지원'],
    ['11', 'A5_2_2', '필요한 정책', '2', '1', '비용 지원'],
    ['12', 'A6_2', '보호 수준(0~10점)', '', '0', '매우 약함'],
  ],
  'Variable Labels': [['Variable Labels', ''], ['A1_6', "' 기업형태'"]],
  'Value Labels': [
    ['Value Labels', ''],
    ['/A1_6', ''], ['', "1 '독립기업'"], ['', "2 '계열사'"],
    ['/A2_1_1 A2_1_2 A2_1_3', ''], ['', "1 '특허권'"], ['', "2 '실용신안권'"], ['', "3 '디자인권'"],
    ['/A4_1A_1 A4_1A_2', ''], ['', "1 '알고 있다'"], ['', "2 '모른다'"],
    ['/A5_2_1 A5_2_2', ''], ['', "1 '비용 지원'"], ['', "2 '인력 지원'"], ['', "3 '정보 제공'"],
  ],
};
// 복수응답은 선택하면 보기 번호, 안 하면 빈칸. DE1은 코드북에 없고 옆 열에 이름이 있음
const int64Data = TG.prepareData([
  ['SEQ', 'LOI(s)', '주체분류', 'DE1', 'WT', 'A1_1', 'A1_3', 'A1_6', 'A2_1_1', 'A2_1_2', 'A2_1_3', 'A4_1A_1', 'A4_1A_2', 'A5_2_1', 'A5_2_2', 'A6_2'],
  [1, 300, '01.대기업', 1, 2, '가', 2001, 1, 1, null, 3, 1, 2, 1, 2, 5],
  [2, 310, '02.중소기업', 2, 1, '나', 2010, 2, 1, 2, null, 2, 2, 2, 3, 8],
  [3, 320, '02.중소기업', 2, 1, '다', 2015, 1, null, null, 3, 1, 1, 1, 3, 10],
  [4, 330, '01.대기업', 1, 1, '라', 2020, 2, 1, 2, 3, 2, 1, 3, 1, 0],
]);
const int64Cb = TG.parseCodebookBook(int64Book, { format: 'auto' });
const int64 = TG.buildItems(int64Cb, int64Data).items;
const iv = (v) => int64.find((i) => i.vars[0] === v);
test('INT64 형식 자동 감지', () => assert.strictEqual(int64Cb.format, 'int64'));
test('INT64 문항 유형', () => {
  assert.strictEqual(iv('A1_1').kind, 'exclude');
  assert.strictEqual(iv('A1_3').kind, 'numeric');
  assert.strictEqual(iv('A1_6').kind, 'single');
  assert.strictEqual(iv('A2_1_1').kind, 'multi01');
  assert.deepStrictEqual(iv('A2_1_1').itemLabels, ['특허권', '실용신안권', '디자인권']);
  assert.strictEqual(iv('A4_1A_1').kind, 'singleset');
  assert.deepStrictEqual(iv('A4_1A_1').itemLabels, ['사업A', '사업B']);
  assert.strictEqual(iv('A5_2_1').kind, 'rank');
  assert.strictEqual(iv('A5_2_1').title, '필요한 정책');
  assert.strictEqual(iv('A6_2').kind, 'scale');
  assert.strictEqual(iv('A6_2').codes[1].label, '1점');
  assert.strictEqual(iv('SEQ').kind, 'exclude');
  assert.strictEqual(iv('LOI(s)').kind, 'exclude');
  assert.strictEqual(iv('WT').kind, 'weight');
});
test('코드북에 없는 배너 변수의 보기 이름을 옆 열에서 찾기', () => {
  const de1 = iv('DE1');
  assert.strictEqual(de1.kind, 'single');
  assert.strictEqual(de1.include, false); // 표는 기본 제외, 배너로는 사용 가능
  assert.deepStrictEqual(de1.codes, [{ code: 1, label: '대기업' }, { code: 2, label: '중소기업' }]);
});
test('번호/빈칸 방식 복수응답 계산', () => {
  const r = TG.computeTables([iv('A2_1_1')], int64Data, { banners: [{ var: 'DE1', label: '규모', codes: iv('DE1').codes }] });
  const t = r.tables[0];
  assert.strictEqual(t.rows[0].n, 4);
  near(t.rows[0].values[0], 75, '특허권');
  near(t.rows[0].values[1], 50, '실용신안권');
  near(t.rows[0].values[2], 75, '디자인권');
  assert.strictEqual(t.rows[1].label, '대기업');
  near(t.rows[1].values[1], 50, '대기업 실용신안권');
});
test('단일응답 묶음 요약표', () => {
  const r = TG.computeTables([iv('A4_1A_1')], int64Data, {});
  const sum = r.tables.find((t) => /요약/.test(t.title));
  assert.ok(sum, '요약표 없음');
  near(sum.rows[0].values[0], 50, '사업A 인지율');
});

console.log('G-CAII 형식');
// 사내 G-CAII 코드북과 같은 구조의 작은 가상 예시
const gBook = {
  '0000_질문지': [
    ['문항ID', '문항타입', '보기', '로직', 'page'],
    ['Q1', '단수응답형 (SIG)', 'SQ1] 활용 수준', '', 1], ['', '', '1]', '', ''], ['', '', '1.높음', '', ''], ['', '', '2.낮음', '', ''],
    ['Q2', '복수응답형 (MTP)', 'B2] 수집 경로 (중복응답)', '', 2], ['', '', '1.포털', '', ''], ['', '', '2.기관', '', ''], ['', '', '9997.기타', '', ''],
    ['Q3', '척도형-복수응답형 (MX3)', 'B5] 데이터 활용 경험과 계획', '', 3], ['', '', '1]교통 데이터', '', ''], ['', '', '2]화재 데이터', '', ''], ['', '', '1.활용 경험있음', '', ''], ['', '', '2.향후 활용계획 있음', '', ''],
    ['Q4', '척도형-멀티형 (MX7)', 'D3] 제도를 알고 계십니까? 이용한 경험이 있으십니까?', '', 4], ['', '', '1]제도A', '', ''], ['', '', '2]제도B', '', ''], ['', '', '1.예', '', ''], ['', '', '2.아니오', '', ''], ['', '', '3.예', '', ''], ['', '', '4.아니오', '', ''],
    ['Q5', '순위형 (RNK)', 'D4] 필요한 정보 제공 방식', '', 5], ['', '', '1.포털', '', ''], ['', '', '2.메일', '', ''], ['', '', '3.설명회', '', ''],
    ['Q6', '숫자형 (NUM)', 'C2] 매출액', '', 6], ['', '', '1.총 매출액', '', ''],
    ['Q7', '자기기입형 (OPN)', 'A1] 기업명', '', 7],
    ['Q8', '단수응답형 (SIG)', 'B5-1] 세부 활용 여부', '', 8], ['', '', '1.예', '', ''], ['', '', '2.아니오', '', ''],
  ],
  'Variable Labels': [
    ['Variable Labels', '', '문구'],
    ['SQ1', '활용 수준', null], ['A1O1', '기업명 - ', null],
    ['B2M1', '수집 경로 (중복응답) - 포털', null], ['B2M2', '수집 경로 (중복응답) - 기관', null], ['B2M9997', '수집 경로 (중복응답) - 기타', ')'],
    ['B5MT1M1', '데이터 활용 경험과 계획 - 교통 데이터', null], ['B5MT1M2', '데이터 활용 경험과 계획 - 교통 데이터', null],
    ['B5MT2M1', '데이터 활용 경험과 계획 - 화재 데이터', null], ['B5MT2M2', '데이터 활용 경험과 계획 - 화재 데이터', null],
    ['D3C1MT1', '제도 - 제도A', null], ['D3C2MT1', '제도 - 제도B', null], ['D3C1MT2', '제도 - 제도A', null], ['D3C2MT2', '제도 - 제도B', null],
    ['D4R1', '필요한 정보 제공 방식 - 1순위', null], ['D4R2', '필요한 정보 제공 방식 - 2순위', null], ['D4R9997', '필요한 정보 제공 방식 - 기타', ')'],
    ['C2N1', '매출액 - 총 매출액', '백만원'], ['B5K1', '세부 활용 여부', null],
  ],
  'Value Labels': [
    ['Value Labels', ''],
    ['/SQ1', ''], ['', "1.'높음'"], ['', "2.'낮음'"],
    ['/B2M1 B2M2 B2M9997', ''], ['', "1.'포털'"], ['', "2.'기관'"], ['', "9997.'기타'"],
    ['/B5MT1M1 B5MT1M2 B5MT2M1 B5MT2M2', ''], ['', "1.'활용 경험있음'"], ['', "2.'향후 활용계획 있음'"],
    ['/D3C1MT1 D3C2MT1 D3C1MT2 D3C2MT2 D3C1MT3', ''], ['', "1.'예'"], ['', "2.'아니오'"], ['', "3.'예'"], ['', "4.'아니오'"],
    ['/D4R1 D4R2', ''], ['', "1.'포털'"], ['', "2.'메일'"], ['', "3.'설명회'"],
    ['/B5K1', ''], ['', "1.'예'"], ['', "2.'아니오'"],
  ],
};
const gData = TG.prepareData([
  ['user_id', 'begin_dt', 'data_지역', 'SQ1', 'A1O1', 'B2M1', 'B2M2', 'B2M9997', 'B2M9997_TXT', 'B5MT1M1', 'B5MT1M2', 'B5MT2M1', 'B5MT2M2', 'D3C1MT1', 'D3C2MT1', 'D3C1MT2', 'D3C2MT2', 'D4R1', 'D4R2', 'D4R9997', 'C2N1', 'B5K1'],
  [11, '2026-01-01', '서울', 1, '가', 1, null, null, null, 1, null, null, 2, 1, 2, 3, null, 1, 2, null, 100, 1],
  [12, '2026-01-01', '경기', 2, '나', 1, 2, null, null, null, null, null, null, 2, 2, null, null, 2, 3, null, 300, 2],
  [13, '2026-01-01', '서울', 1, '다', null, 2, 9997, 'API', 1, 2, 1, null, 1, 1, 4, 3, 3, 1, null, 999999, 1],
  [14, '2026-01-01', '-', 2, '라', 1, null, null, null, null, 2, null, null, 2, 1, null, 4, 1, 3, null, 200, 2],
]);
const gCb = TG.parseCodebookBook(gBook, { format: 'auto' });
const gItems = TG.buildItems(gCb, gData).items;
const gv = (v) => gItems.find((i) => i.vars[0] === v);
test('G-CAII 형식 자동 감지 (INT64로 오인하지 않음)', () => assert.strictEqual(gCb.format, 'gcaii'));
test('G-CAII 변수명 규칙으로 유형 인식', () => {
  assert.strictEqual(gv('SQ1').kind, 'single');
  assert.strictEqual(gv('SQ1').title, 'SQ1. 활용 수준');
  assert.strictEqual(gv('A1O1').kind, 'exclude');
  assert.strictEqual(gv('B2M1').kind, 'multi01');
  assert.deepStrictEqual(gv('B2M1').itemLabels, ['포털', '기관', '기타']);
  assert.strictEqual(gv('B2M9997_TXT').kind, 'exclude');
  assert.strictEqual(gv('D4R1').kind, 'rank');
  assert.deepStrictEqual(gv('D4R1').vars, ['D4R1', 'D4R2']);
  assert.strictEqual(gv('C2N1').kind, 'numeric');
  assert.ok(/단위: 백만원/.test(gv('C2N1').title));
  assert.strictEqual(gv('B5K1').title, 'B5-1. 세부 활용 여부');
  assert.strictEqual(gv('user_id').kind, 'exclude');
});
test('행렬 복수응답(MX3)은 열별로 묶고 전체가 Base', () => {
  const m1 = gv('B5MT1M1');
  assert.deepStrictEqual(m1.vars, ['B5MT1M1', 'B5MT2M1']);
  assert.deepStrictEqual(m1.itemLabels, ['교통 데이터', '화재 데이터']);
  assert.ok(/활용 경험있음$/.test(m1.title));
  const t = TG.computeTables([m1], gData, {}).tables[0];
  assert.strictEqual(t.rows[0].n, 4);
  near(t.rows[0].values[0], 50, '교통 활용경험');
});
test('행렬 멀티(MX7)는 열마다 보기와 제목을 나눔', () => {
  const c1 = gv('D3C1MT1');
  const c2 = gv('D3C1MT2');
  assert.deepStrictEqual(c1.codes.map((c) => c.code), [1, 2]);
  assert.deepStrictEqual(c2.codes.map((c) => c.code), [3, 4]);
  assert.strictEqual(c1.title, 'D3. 제도를 알고 계십니까?');
  assert.strictEqual(c2.title, 'D3. 이용한 경험이 있으십니까?');
});
test('숫자 문항의 999999(모름)는 평균에서 제외', () => {
  const t = TG.computeTables([gv('C2N1')], gData, {}).tables[0];
  assert.strictEqual(t.rows[0].n, 3);
  near(t.rows[0].values[0], 200, '평균');
});
test('문자로만 된 지역 열을 배너용 보기로 변환', () => {
  const r = gv('data_지역');
  assert.strictEqual(r.kind, 'single');
  assert.strictEqual(r.title, '지역');
  assert.deepStrictEqual(r.codes.map((c) => c.label), ['경기', '서울']); // '-'는 빈칸 처리
});

test('G-CAII 질문지 시트만 있는 코드북 (Variable/Value Labels 없음)', () => {
  const book = {
    Questionnaire_예시: [
      ['문항ID', '문항타입', '보기', '로직', 'Page'],
      ['Q3', '단순응답형(Single Answer)', 'SQ1] 소속', null, 2], [null, null, '1. 보육교사'], [null, null, '2. 원장'], [null, null, '9997. 기타'],
      ['Q4', '숫자형(Numeric)', 'SQ2] 보육업무 경력', null, 3], [null, null, '1. _____________년'], [null, null, '2. _____________개월'],
      ['Q15', '척도형(Matrix - single)', '4] 변경사항이 얼마나 적절하다고 생각하십니까?', null, 4],
      [null, null, '1] 문서 간소화'], [null, null, '2] 평가주기'],
      [null, null, '1. '], [null, null, null], [null, null, '매우 적절함'], [null, null, 1],
      [null, null, '2. '], [null, null, '적절함'], [null, null, 2],
      [null, null, '3. '], [null, null, '적절하지 않음'], [null, null, 3],
      [null, null, '4. '], [null, null, '전혀 적절하지 않음'], [null, null, 4],
      ['Q28', '단순응답형(Single Answer)', '12-1] 필요한 이유', null, 5], [null, null, '1. 객관성'], [null, null, '2. 신뢰']
      , ['Q36', '순위형(Ranking)', '17-1] 도움 영역 2가지', null, 6], [null, null, '2순위(==)'], [null, null, '1. 보육과정'], [null, null, '2. 상호작용'], [null, null, '3. 안전'],
      ['Q43', '전화번호형(Phone Number)', '핸드폰번호] 번호 입력', null, 7],
    ],
  };
  const d = TG.prepareData([
    ['응답ID', 'SQ1', 'SQ1_9997ET', 'SQ2_1', 'SQ2_2', 'Y4_1', 'Y4_2', 'Y12K1', 'Y17K1_1순위', 'Y17K1_2순위', '핸드폰번호'],
    [1, 1, null, 3, 2, 1, 2, 1, 1, 2, '010'],
    [2, 2, null, 10, 0, 2, 2, 2, 3, 1, null],
    [3, 9997, 7, null, null, 4, 3, 1, 2, 3, null],
  ]);
  const cb = TG.parseCodebookBook(book, { format: 'auto' });
  assert.strictEqual(cb.format, 'gcaii');
  const items = TG.buildItems(cb, d).items;
  const f = (v) => items.find((i) => i.vars[0] === v);
  assert.strictEqual(f('SQ1').kind, 'single');
  assert.strictEqual(f('SQ1').include, true);
  assert.strictEqual(f('SQ1_9997ET').kind, 'exclude');
  assert.strictEqual(f('SQ2_1').kind, 'numeric');
  assert.strictEqual(f('SQ2_1').title, 'SQ2. 보육업무 경력 - 년');
  assert.strictEqual(f('SQ2_2').title, 'SQ2. 보육업무 경력 - 개월');
  const y4 = f('Y4_1'); // 데이터의 Y 접두어는 떼고 질문 4번에 맞춤
  assert.strictEqual(y4.kind, 'scaleset');
  assert.deepStrictEqual(y4.itemLabels, ['문서 간소화', '평가주기']);
  assert.deepStrictEqual(y4.codes.map((c) => c.label), ['매우 적절함', '적절함', '적절하지 않음', '전혀 적절하지 않음']);
  assert.strictEqual(f('Y12K1').title, '12-1. 필요한 이유');
  assert.strictEqual(f('Y17K1_1순위').kind, 'rank');
  assert.deepStrictEqual(f('Y17K1_1순위').vars, ['Y17K1_1순위', 'Y17K1_2순위']);
  assert.strictEqual(f('핸드폰번호').kind, 'exclude');
  // 1=매우 적절함이 긍정 → 긍정(Top2)은 1·2번
  const t = TG.computeTables([y4], d, {}).tables[0];
  const top = t.columns.findIndex((c) => /Top2/.test(c.label));
  near(t.rows[0].values[top], (2 / 3) * 100, 'Top2');
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

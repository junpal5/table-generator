// 계산 엔진 검증 테스트. 실행: npm test
const assert = require('assert');
const path = require('path');
const XLSX = require('xlsx');
const TG = require('../js/core.js');
require('../js/formats.js');
require('../js/derive.js');
require('../js/render.js');
globalThis.JSZip = require('jszip');
require('../js/hwpx-template.js');
require('../js/export-hwpx.js');
require('../js/export-spss.js');

const SAMPLES = path.join(__dirname, '..', 'samples');
const rows = (f) => {
  const wb = XLSX.readFile(path.join(SAMPLES, f));
  return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null });
};

let passed = 0;
const pending = []; // 비동기 테스트(파일 압축 등)는 끝에서 기다림
function test(name, fn) {
  const ok = () => {
    passed++;
    console.log('  ✓ ' + name);
  };
  const fail = (e) => {
    console.error('  ✗ ' + name + '\n    ' + e.message);
    process.exitCode = 1;
  };
  try {
    const r = fn();
    if (r && typeof r.then === 'function') pending.push(r.then(ok, fail));
    else ok();
  } catch (e) {
    fail(e);
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

console.log('보기 정렬');
test('단일응답: 전체 기준 큰 순서, 기타·모름은 맨 뒤, 계는 그대로', () => {
  const d = TG.prepareData([['G', 'A'], [1, 3], [2, 1], [1, 3], [2, 3], [1, 9], [2, 2], [1, 9], [1, 9], [2, 9]]);
  const c = TG.parseCodebook([['변수명', '문항', '코드', '레이블'], ['A', '이용 채널', 1, '매장'], ['', '', 2, '앱'], ['', '', 3, '웹'], ['', '', 9, '기타'],
    ['G', '성별', 1, '남'], ['', '', 2, '여']]);
  const items = TG.buildItems(c, d).items;
  const a = items.find((i) => i.vars[0] === 'A');
  a.sort = true;
  const t = TG.computeTables([a], d, { banners: [{ var: 'G', label: '성별', codes: items.find((i) => i.vars[0] === 'G').codes }] }).tables[0];
  // 기타(4명)가 가장 크지만 맨 뒤, 나머지는 웹(3) > 매장(1) = 앱(1) → 동률은 원래 순서
  assert.deepStrictEqual(t.columns.map((x) => x.label), ['웹', '매장', '앱', '기타', '계']);
  near(t.rows[0].values[0], (3 / 9) * 100, '웹 %');
  near(t.rows[0].values[3], (4 / 9) * 100, '기타 %');
  // 배너 행도 같은 열 순서로 따라감 (남: 3,9,3,9,9 → 웹 2/5)
  near(t.rows[1].values[0], 40, '남자 웹 %');
  assert.ok(t.notes.some((n) => /큰 순서/.test(n)));
});
test('정렬하지 않으면 코드북 순서 유지', () => {
  const t = plain.tables.find((x) => x.title === 'SQ3. 거주 지역');
  assert.strictEqual(t.columns[0].label, '서울');
});
test('복수응답·순위는 정렬, 척도·수치는 정렬 대상 아님', () => {
  const multi = items.find((i) => i.vars[0] === 'Q3_1');
  const rank = items.find((i) => i.vars[0] === 'Q4_1');
  assert.ok(TG.canSort(multi) && TG.canSort(rank));
  assert.ok(!TG.canSort(items.find((i) => i.vars[0] === 'Q1')));
  assert.ok(!TG.canSort(items.find((i) => i.vars[0] === 'Q2_1')));
  assert.ok(!TG.canSort(items.find((i) => i.vars[0] === 'Q5')));
  const r = TG.computeTables([{ ...multi, sort: true }, { ...rank, sort: true }], data, {});
  r.tables.forEach((t) => {
    const vals = t.columns.map((c, i) => [c, t.rows[0].values[i]]).filter(([c]) => !c.isSum && !/기타/.test(c.label)).map(([, v]) => v);
    for (let i = 1; i < vals.length; i++) assert.ok(vals[i - 1] >= vals[i], `${t.title}: 내림차순이 아님`);
  });
  const rank2 = r.tables.find((t) => /1\+2순위/.test(t.title));
  assert.strictEqual(rank2.columns[rank2.columns.length - 1].label, '기타');
});
test('단일응답 묶음은 요약표만 정렬하고 개별 표는 보기 순서 유지', () => {
  const r = TG.computeTables([{ ...iv('A4_1A_1'), sort: true }], int64Data, {});
  const each = r.tables.filter((t) => t.kind !== 'summary');
  each.forEach((t) => assert.strictEqual(t.columns[0].label, '알고 있다'));
  const sum = r.tables.find((t) => t.kind === 'summary');
  assert.deepStrictEqual(sum.columns.map((c) => c.label), ['사업A', '사업B']); // 50% = 50% → 원래 순서
  assert.ok(sum.sorted);
});

console.log('한글(HWPX) 저장');
const hwpxOpts = { banners: [{ var: 'SQ1', label: '성별', codes: items.find((i) => i.vars[0] === 'SQ1').codes }], decimals: 1, orientation: 'row' };
const hwpxRes = TG.computeTables(items.filter((i) => ['SQ1', 'Q3_1'].includes(i.vars[0])), data, hwpxOpts);
test('HWPX 본문에 표 제목·표·주석이 들어감', () => {
  const f = TG.buildHwpxFiles(hwpxRes, hwpxOpts);
  assert.strictEqual(f.mimetype, 'application/hwp+zip');
  const sec = f['Contents/section0.xml'];
  assert.strictEqual((sec.match(/<hp:tbl /g) || []).length, hwpxRes.tables.length);
  assert.ok(sec.includes('표 1. SQ1. 귀하의 성별은 무엇입니까?'));
  assert.ok(sec.includes('※ Base: 해당 문항 응답자'));
  // "구분"은 2칸 병합, 배너 그룹 "성별"은 2행 병합, 덮인 칸은 적지 않음
  assert.ok(/<hp:t>구분<\/hp:t>.*?<hp:cellSpan colSpan="2" rowSpan="1"\/>/.test(sec));
  assert.ok(/<hp:t>성별<\/hp:t>.*?<hp:cellSpan colSpan="1" rowSpan="2"\/>/.test(sec));
  const firstTbl = sec.slice(sec.indexOf('<hp:tbl '), sec.indexOf('</hp:tbl>'));
  const rows = firstTbl.split('<hp:tr>').slice(1).map((r) => (r.match(/<hp:tc /g) || []).length);
  assert.deepStrictEqual(rows, [5, 5, 6, 5]); // 머리글·전체(구분 2칸 병합) 5칸, 남자 행은 성별 포함 6칸, 여자 행은 성별에 덮여 5칸
  // 열 너비 합 = 본문 폭
  const widths = firstTbl.split('<hp:tr>')[2].match(/cellSz width="(\d+)"/g).map((x) => Number(x.match(/\d+/)[0]));
  assert.strictEqual(widths.reduce((a, b) => a + b, 0), TG.HWPX_PAGE.TEXT_WIDTH);
});
test('HWPX: 한 쪽에 표 하나 (표 제목마다 쪽 나눔)', () => {
  const sec = TG.buildHwpxFiles(hwpxRes, hwpxOpts)['Contents/section0.xml'];
  const breaks = sec.match(/<hp:p [^>]*pageBreak="1"[^>]*>(<hp:run [^>]*>)<hp:t>표 \d+\./g) || [];
  assert.strictEqual(breaks.length, hwpxRes.tables.length);
  assert.strictEqual((sec.match(/pageBreak="1"/g) || []).length, hwpxRes.tables.length);
});
test('HWPX 표는 글자처럼 취급하지 않고, 쪽 경계에서 셀 단위로 나누며 제목 줄 반복', () => {
  const sec = TG.buildHwpxFiles(hwpxRes, hwpxOpts)['Contents/section0.xml'];
  const tbls = sec.match(/<hp:tbl [^>]*>/g);
  tbls.forEach((t) => assert.ok(/pageBreak="CELL"/.test(t) && /repeatHeader="1"/.test(t) && /textWrap="TOP_AND_BOTTOM"/.test(t)));
  assert.ok(!/treatAsChar="1"/.test(sec));
  const head = TG.buildHwpxFiles(hwpxRes, hwpxOpts)['Contents/header.xml'];
  assert.ok(!head.includes('${'), '틀 문자열이 그대로 남음');
  assert.strictEqual((head.match(/keepWithNext="1"/g) || []).length, 1); // 표 제목 문단
  assert.strictEqual((sec.match(/treatAsChar="0"/g) || []).length, hwpxRes.tables.length);
  // 표 높이 = 줄 높이 합
  const first = sec.slice(sec.indexOf('<hp:tbl '), sec.indexOf('</hp:tbl>'));
  const h = Number(/<hp:sz width="\d+" widthRelTo="ABSOLUTE" height="(\d+)"/.exec(first)[1]);
  const rowH = first.split('<hp:tr>').slice(1).map((r) => Number(/cellSz width="\d+" height="(\d+)"/.exec(r.split('rowSpan="1"')[1] ? r.slice(r.indexOf('rowSpan="1"')) : r)[1]));
  assert.strictEqual(h, rowH.reduce((a, b) => a + b, 0));
});
test('HWPX: 긴 표는 글자를 줄여 한 쪽에 맞춤, 짧은 표는 기본 크기', () => {
  const ids = { sizes: [8.5, 8, 7.5, 7, 6.5, 6].map((size, i) => ({ size, cell: i * 2, bold: i * 2 + 1 })) };
  const short = TG.hwpxLayout(hwpxRes.tables[0], hwpxOpts, ids);
  assert.ok(short.fits && short.z.size === 8.5);
  // 보기 60개짜리 배너형 표(행이 많음)
  const big = { no: 9, title: '긴 표', kind: 'single', notes: ['Base: 전체'], columns: Array.from({ length: 60 }, (_, i) => ({ label: '보기 ' + (i + 1), fmt: 'pct' })),
    rows: [{ group: '전체', label: '전체', isTotal: true, n: 100, wn: 100, values: Array.from({ length: 60 }, () => 1.6) }] };
  const L = TG.hwpxLayout(big, { ...hwpxOpts, orientation: 'col' }, ids);
  assert.ok(L.fits, `높이 ${L.total} > ${TG.HWPX_PAGE.TEXT_HEIGHT}`);
  assert.ok(L.z.size < 8.5);
  // 너무 길면 가장 작은 글자로 두고 다음 쪽으로 이어짐
  const huge = { ...big, columns: Array.from({ length: 120 }, (_, i) => ({ label: '보기 ' + (i + 1), fmt: 'pct' })), rows: [{ ...big.rows[0], values: Array.from({ length: 120 }, () => 1) }] };
  const H = TG.hwpxLayout(huge, { ...hwpxOpts, orientation: 'col' }, ids);
  assert.ok(!H.fits && H.z.size === 6);
});
test('HWPX 서식 목록 개수(itemCnt)가 실제 항목 수와 같음', () => {
  const h = TG.buildHwpxFiles(hwpxRes, hwpxOpts)['Contents/header.xml'];
  [['borderFills', 'borderFill'], ['charProperties', 'charPr'], ['paraProperties', 'paraPr']].forEach(([list, item]) => {
    const cnt = Number(new RegExp(`<hh:${list} itemCnt="(\\d+)"`).exec(h)[1]);
    assert.strictEqual((h.match(new RegExp(`<hh:${item} `, 'g')) || []).length, cnt, list);
  });
});
test('HWPX 압축 파일: mimetype이 맨 앞·무압축', async () => {
  const bytes = await TG.buildHwpx(hwpxRes, hwpxOpts);
  // ZIP 첫 항목 이름과 압축 방식(0 = 저장만)
  const nameLen = bytes[26] | (bytes[27] << 8);
  const name = Buffer.from(bytes.slice(30, 30 + nameLen)).toString();
  assert.strictEqual(name, 'mimetype');
  assert.strictEqual(bytes[8] | (bytes[9] << 8), 0);
  const zip = await globalThis.JSZip.loadAsync(bytes);
  assert.ok(zip.file('Contents/section0.xml') && zip.file('Contents/header.xml') && zip.file('Contents/content.hpf'));
});

console.log('11점 척도 묶음');
test('11점 척도(0~10점)는 0~3 / 4~6 / 7~10으로 묶음 (화면·신택스 같은 기준)', () => {
  const t = TG.computeTables([iv('A6_2')], int64Data, {}).tables[0];
  const labels = t.columns.map((c) => c.label);
  assert.ok(labels.includes('긍정(7~10점)') && labels.includes('부정(0~3점)'), labels.join(','));
  // 값 5, 8, 10, 0 → 긍정(7~10) 2명, 부정(0~3) 1명
  near(t.rows[0].values[labels.indexOf('긍정(7~10점)')], 50, '긍정');
  near(t.rows[0].values[labels.indexOf('부정(0~3점)')], 25, '부정');
  const s = TG.buildSyntax(TG.computeTables([iv('A6_2')], int64Data, {}), [iv('A6_2')], int64Data, {});
  assert.ok(s.includes('RECODE G_A6_2(0 1 2 3=21)(4 5 6=22)(7 8 9 10=23)(ELSE=SYSMIS).'));
  assert.ok(s.includes("21 '⊙ 부정(0~3점)'") && s.includes("22 '⊙ 보통(4~6점)'") && s.includes("23 '⊙ 긍정(7~10점)'"));
});
test('5점 척도는 그대로 Top2/Bottom2', () => {
  const t = TG.computeTables([items.find((i) => i.vars[0] === 'Q1')], data, {}).tables[0];
  assert.ok(t.columns.some((c) => c.label === '긍정(Top2)') && t.columns.some((c) => c.label === '부정(Bottom2)'));
});

console.log('가공 변수');
{
  // 설립연도(9999 = 모름), 지역(1~4), 성별
  const rows = [['ID', 'YEAR', 'REG', 'SEX'], [1, 2020, 1, 1], [2, 2015, 2, 2], [3, 2010, 3, 1], [4, 2024, 4, 2], [5, 9999, 1, 1], [6, 2000, null, 2]];
  const cb = TG.parseCodebook([
    ['변수명', '문항', '코드', '레이블'],
    ['YEAR', '설립연도', null, null],
    ['REG', '지역', 1, '서울'], [null, null, 2, '경기'], [null, null, 3, '부산'], [null, null, 4, '경남'],
    ['SEX', '성별', 1, '남'], [null, null, 2, '여'],
  ]);
  const dData = TG.prepareData(rows);
  const built = TG.buildItems(cb, dData);
  const defs = [
    { id: 'a', type: 'compute', name: 'AGE', label: '업력', source: 'YEAR', op: 'k-x', k: '2026' },
    { id: 'b', type: 'bin', name: 'AGE_G', label: '업력 구간', source: 'AGE', bins: '~4 / 5~10 / 11~' },
    { id: 'c', type: 'merge', name: 'AREA', label: '권역', source: 'REG', groups: { 1: '수도권', 2: '수도권', 3: '동남권', 4: '동남권' } },
  ];
  const dv = TG.applyDerived(dData, built.items, defs);
  const di = (name) => dv.items.find((it) => it.vars[0] === name);
  test('구간 문자열 읽기와 겹침 확인', () => {
    const p = TG.parseBins('~4 / 5~9 / 10~');
    assert.deepStrictEqual(p.bins.map((b) => b.label), ['4 이하', '5~9', '10 이상']);
    assert.ok(TG.parseBins('~5 / 5~9').error);
    assert.ok(TG.parseBins('9~5').error);
    assert.ok(TG.parseBins('abc').error);
  });
  test('값 계산: 2026 − 연도, 9999(모름)는 빈 값', () => {
    assert.deepStrictEqual(dData.columns.get(TG.keyOf('AGE')).values, [6, 11, 16, 2, null, 26]);
    assert.strictEqual(di('AGE').kind, 'numeric');
    const t = TG.computeTables([di('AGE')], dData, {}).tables[0];
    near(t.rows[0].values[0], (6 + 11 + 16 + 2 + 26) / 5, '평균');
  });
  test('구간 나누기: 가공 변수를 다시 가공, 표에 구간 분포 + 원래 값 평균', () => {
    assert.deepStrictEqual(dData.columns.get(TG.keyOf('AGE_G')).values, [2, 3, 3, 1, null, 3]);
    const t = TG.computeTables([di('AGE_G')], dData, {}).tables[0];
    assert.deepStrictEqual(t.columns.map((c) => c.label), ['4 이하', '5~10', '11 이상', '계', '평균']);
    near(t.rows[0].values[0], 20, '4 이하');
    near(t.rows[0].values[4], 12.2, '평균');
  });
  test('보기 묶기: 권역으로 묶어 배너로 사용', () => {
    assert.deepStrictEqual(di('AREA').codes.map((c) => c.label), ['수도권', '동남권']);
    const t = TG.computeTables([di('AGE')], dData, { banners: [{ var: 'AREA', label: '권역', codes: di('AREA').codes }] }).tables[0];
    const sudo = t.rows.find((r) => r.label === '수도권');
    assert.strictEqual(sudo.n, 2); // YEAR 9999는 계산에서 빠짐
    near(sudo.values[0], 8.5, '수도권 평균');
  });
  test('다시 적용해도 변수가 겹치지 않고, 고친 표 제목은 유지', () => {
    di('AGE').title = '업력(년)';
    const again = TG.applyDerived(dData, dv.items, defs);
    assert.strictEqual(again.items.filter((it) => it.derived).length, 3);
    assert.strictEqual(dData.order.filter((k) => dData.columns.get(k).derived).length, 3);
    assert.strictEqual(again.items.find((it) => it.vars[0] === 'AGE').title, '업력(년)');
  });
  test('잘못된 설정은 만들지 않고 알려 줌', () => {
    assert.ok(TG.checkDerived({ id: 'x', type: 'compute', name: 'SEX', source: 'YEAR', op: 'k-x', k: '1' }, dData, dv.items, []));
    assert.ok(TG.checkDerived({ id: 'x', type: 'compute', name: '1AB', source: 'YEAR', op: 'k-x', k: '1' }, dData, dv.items, []));
    assert.ok(TG.checkDerived({ id: 'x', type: 'compute', name: 'NEWV', source: 'YEAR', op: 'x/k', k: '0' }, dData, dv.items, []));
    const r = TG.applyDerived(TG.prepareData(rows), TG.buildItems(cb, TG.prepareData(rows)).items, [{ id: 'z', type: 'bin', name: 'Z', source: 'NOPE', bins: '~1' }]);
    assert.ok(r.warnings.length && !r.items.some((it) => it.derived));
  });
  test('SPSS 신택스: 가공 변수 COMPUTE/RECODE와 구간 표의 평균', () => {
    const items2 = TG.applyDerived(dData, dv.items, defs).items;
    const use = items2.filter((it) => ['AGE', 'AGE_G'].includes(it.vars[0]));
    const opts = { banners: [{ var: 'AREA', label: '권역', codes: di('AREA').codes }] };
    const s = TG.buildSyntax(TG.computeTables(use, dData, opts), items2, dData, opts);
    assert.ok(s.includes('IF (NOT(ANY(YEAR,9999))) AGE=2026-YEAR.'), 'compute');
    assert.ok(s.includes('RECODE AGE (LO THRU 4=1)(5 THRU 10=2)(11 THRU HI=3)(ELSE=SYSMIS) INTO AGE_G.'), 'bin');
    assert.ok(s.includes('RECODE REG (1 2=1)(3 4=2)(ELSE=SYSMIS) INTO AREA.'), 'merge');
    assert.ok(s.indexOf('INTO AREA.') < s.indexOf('COMPUTE T_AREA=AREA.'), '가공 변수가 배너보다 먼저');
    assert.ok(s.includes('IF (NOT MISSING(AGE_G)) AA_AGE_G=AGE.'));
    assert.ok(s.includes('BY T2+AGE_G+T3+AA_AGE_G'));
  });
}

console.log('SPSS 신택스');
{
  const sOpts = { banners: [{ var: 'SQ1', label: '성별', codes: items.find((i) => i.vars[0] === 'SQ1').codes }], weightVar: 'wt', decimals: 1 };
  const sRes = TG.computeTables(items, data, sOpts);
  const sps = TG.buildSyntax(sRes, items, data, sOpts);
  test('표마다 TABLE 명령 하나, 번호·제목이 화면과 같음', () => {
    const nos = (sps.match(/^\* NO=(\d+)\.\r?$/gm) || []).map((x) => Number(x.match(/\d+/)[0]));
    assert.deepStrictEqual(nos, sRes.tables.map((t) => t.no));
    assert.strictEqual((sps.match(/^TABLE\r?$/gm) || []).length, sRes.tables.length + 1); // + 응답자 특성
    sRes.tables.forEach((t) => assert.ok(sps.includes(`"표 ${t.no}. `), `표 ${t.no} 제목 없음`));
  });
  test('파일 경로·저장 명령은 넣지 않음, 가중치는 WEIGHT BY', () => {
    assert.ok(!/^\s*(CD|GET FILE|SAVE|XSAVE)\b/im.test(sps));
    assert.ok(!/OUTFILE/i.test(sps));
    assert.ok(/^WEIGHT BY wt\.\r?$/m.test(sps));
    assert.ok(/^EXECUTE\.\r?$/m.test(sps));
  });
  test('사내 형식: 배너 T_/V_ 정의, /PTO·/STA·/CORNER', () => {
    assert.ok(sps.includes('COMPUTE T_SQ1=SQ1.') && sps.includes('COMPUTE V_SQ1=SQ1.'));
    assert.ok(sps.includes("VALUE LABEL T_SQ1 V_SQ1"));
    assert.ok(sps.includes("/PTO=T1'■ 전      체 ■' T2''/FTO=T3'계'"));
    assert.ok(sps.includes("/TAB=T1+T_SQ1 BY T2+SQ2+T3"));
    assert.ok(sps.includes("CPC(SQ2(F4.1)'%':T_SQ1)"));
    assert.ok(sps.includes('/CORNER="BASE=전체"'));
  });
  test('복수응답(0/1)은 /MDG, 순위 종합은 /MRG, 척도는 평균·100점·긍정/부정 묶음', () => {
    assert.ok(sps.includes("/MDG=M''Q3_1 Q3_2 Q3_3 Q3_4 Q3_5 Q3_6"));
    assert.ok(sps.includes("/MRG=M''Q4_1 Q4_2"));
    assert.ok(sps.includes('IF (Q1 GE 1 AND Q1 LE 5) AA_Q1=Q1.'));
    assert.ok(sps.includes('IF (Q1 GE 1 AND Q1 LE 5) Y_Q1=(AA_Q1-1)*100/4.'));
    assert.ok(sps.includes('RECODE G_Q1(1 2=21)(3=22)(4 5=23)(ELSE=SYSMIS).'));
    assert.ok(sps.includes("/MRG=M''Q1 G_Q1"));
  });
  test('번호/빈칸 복수응답은 묶음 전체에 같은 값 레이블(/MRG가 첫 변수 레이블을 씀)', () => {
    const r = TG.computeTables([iv('A2_1_1')], int64Data, {});
    const s2 = TG.buildSyntax(r, [iv('A2_1_1')], int64Data, {});
    assert.ok(/\/A2_1_1 A2_1_2 A2_1_3\r?\n\t1 '특허권'\r?\n\t2 '실용신안권'\r?\n\t3 '디자인권'/.test(s2));
    assert.ok(s2.includes("/MRG=M''A2_1_1 A2_1_2 A2_1_3"));
    assert.ok(/^WEIGHT OFF\.\r?$/m.test(s2));
  });
  test('문자로 된 배너는 RECODE … INTO로 숫자 코드로 바꿈', () => {
    const reg = gv('data_지역');
    const r = TG.computeTables([gv('SQ1')], gData, { banners: [{ var: 'data_지역', label: '지역', codes: reg.codes }] });
    const s3 = TG.buildSyntax(r, [gv('SQ1')], gData, { banners: [{ var: 'data_지역', label: '지역', codes: reg.codes }] });
    assert.ok(s3.includes("RECODE data_지역 ('경기'=1)('서울'=2) INTO T_data_지역."));
  });
  test('인코딩: UTF-8은 BOM+머리말, EUC-KR은 CP949 바이트', () => {
    const u = TG.encodeSyntax('표', 'utf-8');
    assert.deepStrictEqual(Array.from(u.slice(0, 3)), [0xef, 0xbb, 0xbf]);
    const saved = globalThis.cptable;
    globalThis.cptable = require('xlsx/dist/cpexcel.js');
    const e = TG.encodeSyntax('표', 'euc-kr');
    globalThis.cptable = saved;
    const txt = Buffer.from(e).toString('latin1');
    assert.ok(txt.startsWith('* Encoding: EUC-KR.'));
    assert.deepStrictEqual(Array.from(e.slice(-2)), [0xc7, 0xa5]); // '표'
    // EUC-KR에 없는 글자(−, ‧ 등)는 빈 바이트 대신 비슷한 글자로
    globalThis.cptable = require('xlsx/dist/cpexcel.js');
    const e2 = TG.encodeSyntax("'2026 − A1_3' '분쟁‧소송' '× ÷ ·'", 'euc-kr');
    globalThis.cptable = saved;
    assert.ok(!Array.from(e2).includes(0));
    const back = new TextDecoder('euc-kr').decode(e2);
    assert.ok(back.includes("'2026 - A1_3' '분쟁·소송' '× ÷ ·'"), back);
  });
}

console.log('표 모양');
test('보고서형/배너형 격자', () => {
  const t = tbl('SQ1. 귀하의 성별은 무엇입니까?');
  const g1 = TG.toGrid(t, { orientation: 'row' });
  assert.strictEqual(g1.length, 1 + 3);
  const g2 = TG.toGrid(t, { orientation: 'col' });
  assert.strictEqual(g2[0][1].v, '전체');
  assert.strictEqual(g2.length, 2 + 1 + 3); // 머리글 2 + 사례수 + 보기 2 + 계
});

Promise.all(pending).then(() => console.log(`\n${passed}개 통과`));

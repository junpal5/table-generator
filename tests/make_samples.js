// 예시 코드북 / 예시 데이터 / 빈 코드북 양식을 samples/ 폴더에 만듭니다.
// 실행: npm run samples
const ExcelJS = require('exceljs');
const path = require('path');

const OUT = path.join(__dirname, '..', 'samples');

// 결과가 매번 같도록 고정된 난수
let seed = 20260925;
function rand() {
  seed = (seed * 1103515245 + 12345) % 2147483648;
  return seed / 2147483648;
}
function pick(weights) {
  const total = weights.reduce((a, b) => a + b, 0);
  let r = rand() * total;
  for (let i = 0; i < weights.length; i++) {
    r -= weights[i];
    if (r < 0) return i + 1;
  }
  return weights.length;
}

const SCALE5 = [[1, '전혀 그렇지 않다'], [2, '그렇지 않은 편이다'], [3, '보통이다'], [4, '그런 편이다'], [5, '매우 그렇다']];
const SAT5 = [[1, '매우 불만족'], [2, '불만족'], [3, '보통'], [4, '만족'], [5, '매우 만족'], [9, '모름/무응답']];
const CHANNELS = ['인터넷 검색', 'SNS', 'TV·라디오', '지인 추천', '신문·잡지', '옥외광고'];
const REASONS = [[1, '가격'], [2, '품질'], [3, '접근성'], [4, '브랜드'], [5, '서비스'], [6, '디자인'], [7, '기타']];

// [변수명, 문항, 보기 목록, 유형(선택)]
const CODEBOOK = [
  ['ID', '응답자 ID', [], '제외'],
  ['SQ1', 'SQ1. 귀하의 성별은 무엇입니까?', [[1, '남자'], [2, '여자']]],
  ['SQ2', 'SQ2. 귀하의 연령은 어떻게 되십니까?', [[1, '20대'], [2, '30대'], [3, '40대'], [4, '50대'], [5, '60대 이상']]],
  ['SQ3', 'SQ3. 거주 지역', [[1, '서울'], [2, '인천/경기'], [3, '대전/충청'], [4, '광주/전라'], [5, '대구/경북'], [6, '부산/울산/경남'], [7, '강원/제주']]],
  ['Q1', 'Q1. 전반적인 서비스 만족도', SAT5],
  ['Q2_1', 'Q2. 서비스 평가 - 가격이 적절하다', SCALE5],
  ['Q2_2', 'Q2. 서비스 평가 - 품질이 우수하다', SCALE5],
  ['Q2_3', 'Q2. 서비스 평가 - 직원이 친절하다', SCALE5],
  ['Q2_4', 'Q2. 서비스 평가 - 이용이 편리하다', SCALE5],
  ...CHANNELS.map((c, i) => [`Q3_${i + 1}`, `Q3. 서비스를 알게 된 경로(모두 선택) - ${c}`, [[0, '비선택'], [1, '선택']]]),
  ['Q4_1', 'Q4. 서비스 선택 이유(1순위)', REASONS],
  ['Q4_2', 'Q4. 서비스 선택 이유(2순위)', []],
  ['Q5', 'Q5. 최근 1개월 이용 횟수(회)', []],
  ['Q6', 'Q6. (Q1=1,2 응답자) 불만족 이유', [[1, '비싼 가격'], [2, '낮은 품질'], [3, '불친절'], [4, '기타']]],
  ['Q7', 'Q7. 개선 의견(개방형)', [], '제외'],
  ['wt', '가중치', [], '가중치'],
];

async function makeCodebook(file, withData) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('코드북');
  ws.columns = [
    { header: '변수명', width: 10 },
    { header: '문항', width: 46 },
    { header: '코드', width: 8 },
    { header: '레이블', width: 22 },
    { header: '문항유형(선택)', width: 16 },
    { header: '그룹(선택)', width: 12 },
  ];
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE7EEF8' } };
  if (withData) {
    CODEBOOK.forEach(([name, q, codes, type]) => {
      if (!codes.length) ws.addRow([name, q, '', '', type || '']);
      codes.forEach(([code, label], i) => ws.addRow(i === 0 ? [name, q, code, label, type || ''] : ['', '', code, label]));
    });
  } else {
    ws.addRow(['SQ1', 'SQ1. 귀하의 성별은?', 1, '남자']);
    ws.addRow(['', '', 2, '여자']);
    ws.addRow(['Q1', 'Q1. 만족도', 1, '매우 불만족']);
    ws.addRow(['', '', 5, '매우 만족']);
  }
  const help = wb.addWorksheet('작성 방법');
  [
    ['작성 방법'],
    ['1. 한 행에 보기 하나씩 작성합니다. 변수명과 문항은 첫 번째 보기 행에만 적어도 됩니다.'],
    ['2. 보기가 없는 숫자 문항(나이, 횟수 등)은 코드·레이블을 비워 두면 평균표로 만들어집니다.'],
    ['3. 복수응답(0/1)은 Q3_1, Q3_2 …처럼 "문항번호_번호"로 변수명을 지으면 자동으로 묶입니다.'],
    ['4. 순위응답은 Q4_1(1순위), Q4_2(2순위)처럼 작성하고 보기는 첫 번째 변수에만 적어도 됩니다.'],
    ['5. 문항유형은 비워 두면 자동 인식합니다. 직접 지정하려면 단일/척도/복수/순위/복수코드/수치/제외/가중치 중 하나를 적습니다.'],
    ['6. 그룹은 변수명 규칙과 다르게 묶고 싶을 때만 같은 이름(예: Q10)을 적습니다.'],
    ['7. "1=남자, 2=여자"처럼 한 칸에 보기를 모아 적은 코드북도 읽을 수 있습니다.'],
  ].forEach((r) => help.addRow(r));
  help.getColumn(1).width = 110;
  help.getRow(1).font = { bold: true, size: 13 };
  await wb.xlsx.writeFile(file);
}

async function makeData(file) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('data');
  const names = CODEBOOK.map((c) => c[0]);
  ws.addRow(names);
  const N = 500;
  const opinions = ['가격 인하 희망', '앱 개선 필요', '', '', '상담 연결이 빨랐으면', ''];
  for (let i = 1; i <= N; i++) {
    const sex = pick([48, 52]);
    const age = pick([18, 20, 22, 21, 19]);
    const region = pick([19, 26, 11, 10, 10, 15, 9]);
    const lean = age >= 4 ? 0.4 : 0; // 연령이 높을수록 만족도가 약간 높게
    const q1 = rand() < 0.03 ? 9 : pick([5, 10, 30, 40 + lean * 20, 15 + lean * 20]);
    const q2 = [0, 1, 2, 3].map((j) => pick([4, 10, 30 - j * 3, 38, 18 + j * 4]));
    let q3 = CHANNELS.map((_, j) => (rand() < [0.55, 0.35, 0.25, 0.3, 0.1, 0.08][j] ? 1 : 0));
    if (!q3.some(Boolean)) q3[0] = 1;
    const r1 = pick([30, 25, 15, 12, 10, 6, 2]);
    let r2 = r1;
    while (r2 === r1) r2 = pick([20, 20, 20, 15, 15, 8, 2]);
    const hasR2 = rand() < 0.92;
    const q5 = Math.max(0, Math.round(2 + rand() * 8 + (sex === 1 ? 1 : 0)));
    const q6 = q1 === 1 || q1 === 2 ? pick([40, 30, 20, 10]) : null;
    const q7 = opinions[Math.floor(rand() * opinions.length)] || null;
    // 성·연령 가중치 예시(모집단 대비 보정)
    const wt = Number(((sex === 1 ? 1.04 : 0.96) * [1.1, 0.95, 0.9, 1.0, 1.08][age - 1]).toFixed(4));
    ws.addRow([i, sex, age, region, q1, ...q2, ...q3, r1, hasR2 ? r2 : null, q5, q6, q7, wt]);
  }
  await wb.xlsx.writeFile(file);
}

(async () => {
  await makeCodebook(path.join(OUT, '예시_코드북.xlsx'), true);
  await makeCodebook(path.join(OUT, '코드북_양식.xlsx'), false);
  await makeData(path.join(OUT, '예시_데이터.xlsx'));
  console.log('samples/ 폴더에 예시 파일 3개를 만들었습니다.');
})();

// 견본 설문지 생성기 — 지금까지 실제 설문지에서 문제됐던 서식 패턴을 한 파일에 모은다.
// 실제 클라이언트 설문지를 리포지토리에 두지 않고도 파서 회귀를 검증하기 위한 파일.
//   node tests/build_fixture.js
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Tab, Table, TableRow, TableCell,
  WidthType, ShadingType, LevelFormat, AlignmentType,
} = require("docx");

const W = 9360; // 표 전체 폭 (DXA)

function p(text, opts = {}) {
  return new Paragraph({ children: [new TextRun({ text, bold: !!opts.bold })], ...opts.para });
}

function bullet(text) {
  return new Paragraph({ text, numbering: { reference: "opt-bullets", level: 0 } });
}

// 한 단락에 보기 여러 개를 탭으로 나열 (`1) 남자    2) 여자`)
function optionLine(parts) {
  const children = [];
  parts.forEach((t, i) => {
    if (i > 0) children.push(new TextRun({ children: [new Tab(), new Tab()] }));
    children.push(new TextRun({ text: t }));
  });
  return new Paragraph({ children });
}

// 소프트 리턴(shift+enter)으로 두 줄이 된 문항 머리글
function softBreakHeader(first, second) {
  return new Paragraph({
    children: [
      new TextRun({ text: first, bold: true }),
      new TextRun({ text: second, break: 1 }),
    ],
  });
}

function table(rows, colCount) {
  const widths = Array(colCount).fill(Math.floor(W / colCount));
  return new Table({
    columnWidths: widths,
    rows: rows.map((cells, r) =>
      new TableRow({
        children: cells.map((c, i) =>
          new TableCell({
            width: { size: widths[i], type: WidthType.DXA },
            shading: r === 0 ? { type: ShadingType.CLEAR, fill: "F2F2F2" } : undefined,
            children: [new Paragraph({ children: [new TextRun({ text: String(c), bold: r === 0 })] })],
          })),
      })),
  });
}

const body = [
  p("[견본] 파서 회귀 검증용 설문지", { bold: true, para: { alignment: AlignmentType.CENTER } }),
  p("실제 조사에 사용하는 문서가 아닙니다. 서식 패턴만 모아둔 테스트 파일입니다."),

  // 1. 한 단락 탭 구분 보기
  p("SQ1. 귀하의 성별은 무엇입니까? [단수]", { bold: true }),
  optionLine(["1) 남자", "2) 여자"]),

  // 2. 밑줄 숫자 기입 + 하위문항 존재 (섹션 머리글 오인 방지)
  p("SQ2. 귀하의 출생년도는 어떻게 되십니까?", { bold: true }),
  p("_____________년"),
  p("SQ2-1. 연령", { bold: true }),
  optionLine(["1) 20~29세", "2) 30~39세", "3) 40세 이상"]),

  // 3. 보기가 표 안에 여러 줄로 (17개 시·도)
  p("SQ3. 현재 거주하시는 곳은 어디입니까? [단수]", { bold: true }),
  table([
    ["1) 서울", "2) 부산", "3) 대구", "4) 인천", "5) 광주", "6) 대전", "7) 울산"],
    ["8) 경기", "9) 강원", "10) 충북", "11) 충남", "12) 세종", "13) 전북", "14) 전남"],
    ["15) 경북", "16) 경남", "17) 제주", "", "", "", ""],
  ], 7),
  p("[PROG: 지도로 제시]"),

  // 4. 워드 자동 불릿 보기 (텍스트에 기호가 남지 않음) + 하위문항
  p("SQ4. 구매를 시작하신 지 얼마나 되셨습니까? [단수]", { bold: true }),
  bullet("구매 경험 없음"),
  bullet("1개월 미만"),
  bullet("1개월 이상~1년 미만"),
  bullet("1년 이상"),

  // 5. RANGE 표기 → 숫자 문항
  p("SQ4-1. 구매하기 시작하신 지 몇 년 되셨습니까? [직접 입력]", { bold: true }),
  p("(   )년 [RANGE : 1~23]"),

  // 6. 동그라미 보기
  p("Q1. 귀 기관의 유형은 무엇입니까? [1개 선택]", { bold: true }),
  p("① 국공립 어린이집"),
  p("② 보건소"),
  p("③ 공공의료시설(병원)"),
  p("④ 기타 (          )"),

  // 7. 모름/무응답 자동 결측 제안
  p("Q2. 귀하께서는 어느 계층에 속한다고 생각하십니까? [1개 선택]", { bold: true }),
  p("1) 상위"),
  p("2) 중간"),
  p("3) 하위"),
  p("99) 모름/무응답"),

  // 8. 체크박스 + 금액 입력 격자
  p("Q3. 적용된 요소기술과 투자금액을 기입하여 주십시오.", { bold: true }),
  p("[항목별 선택 및 숫자 기입]"),
  table([
    ["기술 항목", "도입 여부", "투자금액(만원)"],
    ["단열(벽체)", "□", "약 (          )만원"],
    ["창호", "□", "약 (          )만원"],
    ["조명", "□", "약 (          )만원"],
  ], 3),
  p("[Range: 투자금액 1~9,999,999만원]"),

  // 9. 격자 + 속성 번호 결번(5 없음) + 코드 칸이 `1)` 형태
  p("Q4. 다음 각 문장에 얼마나 동의하십니까? [행별 1개 선택]", { bold: true }),
  table([
    ["속성", "전혀 동의 안함", "동의 안함", "보통", "동의함", "매우 동의함"],
    ["1. 눈길을 끈다", "1)", "2)", "3)", "4)", "5)"],
    ["2. 다른 제품과 다르다", "1)", "2)", "3)", "4)", "5)"],
    ["4. 구입하고 싶다", "1)", "2)", "3)", "4)", "5)"],
    ["6. 신뢰가 간다", "1)", "2)", "3)", "4)", "5)"],
  ], 6),

  // 10. 항목 열 없는 척도표 (첫 척도점 누락 방지) + 코드 행
  p("Q5. 전반적으로 얼마나 만족하십니까? [1개 선택]", { bold: true }),
  table([
    ["매우 불만족", "불만족", "보통", "만족", "매우 만족"],
    ["1", "2", "3", "4", "5"],
  ], 5),

  // 11. 코드가 전혀 없는 척도표 → 순차 부여 + 확인필요
  p("Q6. 사업 운영에 대한 만족은 어떠십니까? [1개 선택]", { bold: true }),
  table([
    ["매우 불만족", "불만족", "보통", "만족", "매우 만족"],
    ["", "", "", "", ""],
  ], 5),

  // 12. 소프트 리턴으로 두 줄이 된 머리글
  softBreakHeader("Q7. 기대했던 것과 비교할 때, ",
    "제품 외관이 어떻다고 생각하십니까? [1개 선택]"),
  table([
    ["훨씬 나쁘다", "약간 나쁘다", "보통", "약간 좋다", "훨씬 좋다"],
    ["1", "2", "3", "4", "5"],
  ], 5),

  // 13. 복수응답
  p("Q8. 최근 1년 이내 경험한 종류를 모두 선택해 주십시오. [복수]", { bold: true }),
  optionLine(["1) 카지노", "2) 경마", "3) 복권"]),
  optionLine(["4) 기타(       )", "5) 없음"]),

  // 14. 순위 표 + 보기 목록 → 숫자 순위형
  p("Q9. 이용하시는 이유를 순위대로 선택해 주세요. [1순위 필수, 2,3순위 선택]", { bold: true }),
  table([["1순위", "", "2순위", "", "3순위", ""]], 6),
  optionLine(["1) 가까워서", "2) 익숙해서"]),
  optionLine(["3) 정보를 얻기 쉬워서", "4) 기타(   )"]),

  // 15. 문자 접미 머리글 + 전/후 숫자 입력 격자
  p("Q10-a. 리모델링 전·후 요금을 기입해 주십시오.", { bold: true }),
  p("[항목별 숫자 기입]"),
  table([
    ["구분", "리모델링 전", "리모델링 후"],
    ["월평균 전기요금 (만원)", "(          ) 만원", "(          ) 만원"],
    ["월평균 사용량 (kWh)", "(          ) kWh", "(          ) kWh"],
  ], 3),

  // 16. Code1..CodeN 보기 표 + 최대 N개 → 복수응답
  p("Q11. 가장 마음에 드는 재료는 무엇인가요? [최소 1개, 최대 2개]", { bold: true }),
  table([
    ["", "Code1", "Code2", "Code3"],
    ["식재료", "치즈", "토마토", "올리브"],
  ], 4),

  // 17. "최대 3순위 선택" 문구 → 순위형
  p("Q12. 활용 분야를 최대 3순위 선택해 주십시오.", { bold: true }),
  p("① 시설 유지보수"),
  p("② 서비스 향상"),
  p("③ 기자재 구입"),
  p("④ 인건비 충당"),

  // 18. 블록 머리글(하위문항만 존재) → 변수 생성 안 함
  p("Q13. 다음 항목에 대해 평소 생각하시는 바를 알려주십시오.", { bold: true }),
  p("Q13-1. 우리 사회는 공정하다고 생각한다.", { bold: true }),
  table([
    ["전혀 그렇지 않다 (1)", "(2)", "(3)", "(4)", "매우 그렇다 (5)"],
    ["1", "2", "3", "4", "5"],
  ], 5),
  p("Q13-2. 우리 사회는 투명하다고 생각한다.", { bold: true }),
  table([
    ["전혀 그렇지 않다 (1)", "(2)", "(3)", "(4)", "매우 그렇다 (5)"],
    ["1", "2", "3", "4", "5"],
  ], 5),

  // 19. 드롭다운 — 보기 목록이 문서에 없음
  p("Q14. 거주 지역을 선택하여 주십시오. [드롭박스 제시]", { bold: true }),
  p("[PROG: 시/군/구까지 선택할 수 있는 드롭박스 제시]"),

  // 20. 주관식
  p("Q15. 개선이 필요한 점을 자유롭게 적어주십시오. [직접 기입]", { bold: true }),
  table([[""]], 1),

  // 21. 문항 서두가 긴 격자 (라벨 절단 시 항목이 사라지는 문제)
  p("Q16. 다음은 귀 기관이 최근 1년간 이용한 매체별 이용 빈도를 묻는 문항입니다. "
    + "각 매체에 대해 평소 이용하시는 빈도를 하나씩 선택해 주시기 바랍니다. "
    + "정확한 수치를 모르는 경우 가장 가까운 값을 선택해 주십시오. [행별 1개 선택]", { bold: true }),
  table([
    ["언론사", "전혀 이용 안함", "월 1회", "주 1회", "매일"],
    ["1. KBS", "1", "2", "3", "4"],
    ["2. MBC", "1", "2", "3", "4"],
  ], 5),

  // 22. 격자 항목이 자기 변수명을 갖는 경우
  p("EQD. 다음 각 문장에 동의하는 정도를 응답해 주십시오. [행별 1개 선택]", { bold: true }),
  table([
    ["분배 비례성 (EQD)", "전혀 그렇지 않다 1", "2", "3", "4", "매우 그렇다 5"],
    ["1) EQD1. 노력한 만큼 보상받는다", "1", "2", "3", "4", "5"],
    ["2) EQD2. 기회가 공평하게 주어진다", "1", "2", "3", "4", "5"],
  ], 6),

  // 23. 빈칸 기입 양식표 → 문자형
  p("Com1. 응답 기관 정보를 기입해 주십시오. [직접 기입]", { bold: true }),
  table([["기관명", ""], ["담당자 연락처", ""], ["담당자 이메일", ""]], 2),

  // 24. 마침표 없는 머리글 + 공백 하나로 구분된 보기
  p("Com1-2 (Com1 응답자만) 다음 중 가장 선호하는 개선안은 무엇입니까? [단수]", { bold: true }),
  p("1) 개선안A 2) 개선안B 3) 개선안C"),

  // 25. 구분자가 `탭+공백`, 그리고 줄 끝 지시문에 보기 표시가 들어간 경우
  //     (둘 다 보기가 앞 라벨에 흡수되어 조용히 사라지던 패턴)
  p("Q17. 새롭게 도입되었으면 하는 종목을 모두 선택해 주십시오. [복수]", { bold: true }),
  new Paragraph({
    children: [
      new TextRun({ text: "1) 당구" }),
      new TextRun({ children: [new Tab(), new Tab()] }),
      new TextRun({ text: " 2) 탁구" }),          // 탭 뒤에 공백
      new TextRun({ children: [new Tab()] }),
      new TextRun({ text: "3) e스포츠 4) 핸드볼" }), // 공백 하나로만 구분
      new TextRun({ text: " [PROG : 4) 기타 제외 보기 rotation]" }), // 지시문 안의 보기 표시
    ],
  }),

  // 26. 양 끝에만 라벨이 있는 9점 척도 (중간 코드가 사라지면 안 됨)
  p("Q18. 다음 의견에 얼마나 동의하십니까? [1개 선택]", { bold: true }),
  table([
    ["전혀 그렇지 않다", "", "", "", "", "", "", "", "매우 그렇다"],
    ["1", "2", "3", "4", "5", "6", "7", "8", "9"],
  ], 9),

  // 27. DP 작업지시로 생성되는 변수
  p("[DP: 최초 오답여부: IN1_FAIL 변수 만들어주세요.]"),
];

const doc = new Document({
  numbering: {
    config: [{
      reference: "opt-bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children: body }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.mkdirSync("tests", { recursive: true });
  fs.writeFileSync("tests/fixture_patterns.docx", buf);
  console.log("생성: tests/fixture_patterns.docx", buf.length, "bytes");
});

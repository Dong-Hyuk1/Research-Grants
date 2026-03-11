const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "PSR 과제";
pres.title = "PSR 데이터 AI 추론 매핑";

const slide = pres.addSlide();
slide.background = { color: "FFFFFF" };

// Colors
const navy = "1E3A5F";
const blue = "4472C4";
const lightBlue = "D6E4F0";
const green = "2E7D32";
const lightGreen = "E8F5E9";
const orange = "E65100";
const lightOrange = "FFF3E0";
const purple = "6A1B9A";
const lightPurple = "F3E5F5";
const gray = "546E7A";
const darkGray = "263238";

// Title
slide.addText("PSR 데이터 → AI 추론 매핑 체계", {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 20, fontFace: "맑은 고딕", bold: true, color: navy, margin: 0
});

// === LEFT COLUMN: PSR 데이터 (웨어러블) ===
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 0.3, y: 0.9, w: 2.6, h: 0.45,
  fill: { color: navy }, rectRadius: 0.08
});
slide.addText("PSR 데이터 (웨어러블)", {
  x: 0.3, y: 0.9, w: 2.6, h: 0.45,
  fontSize: 12, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// Sensor boxes
const sensors = [
  { label: "Smart Watch", sub: "가속도/자이로, 심박, BIA", y: 1.5 },
  { label: "S-Patch (FDA)", sub: "HRR, HRV, 파워 출력", y: 2.15 },
  { label: "Bio Armour", sub: "ROM, 근력(Arm Curl, STS)", y: 2.8 },
  { label: "Mobile App", sub: "Pose Detection (33 키포인트)", y: 3.45 },
];

sensors.forEach(s => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.3, y: s.y, w: 2.6, h: 0.55,
    fill: { color: lightBlue }, line: { color: blue, width: 1.5 }, rectRadius: 0.06
  });
  slide.addText([
    { text: s.label, options: { bold: true, fontSize: 10, color: navy, breakLine: true } },
    { text: s.sub, options: { fontSize: 8, color: gray } }
  ], {
    x: 0.4, y: s.y, w: 2.4, h: 0.55,
    fontFace: "맑은 고딕", valign: "middle", margin: 0
  });
});

// === MIDDLE COLUMN: 실증 종목 5종 ===
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.5, y: 0.9, w: 2.8, h: 0.45,
  fill: { color: navy }, rectRadius: 0.08
});
slide.addText("실증 종목 5종", {
  x: 3.5, y: 0.9, w: 2.8, h: 0.45,
  fontSize: 12, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// 역학적 분석 group
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.4, y: 1.5, w: 3.0, h: 1.2,
  fill: { color: lightGreen }, line: { color: green, width: 1 }, rectRadius: 0.06
});
slide.addText("역학적 분석", {
  x: 3.5, y: 1.5, w: 1.2, h: 0.3,
  fontSize: 8, fontFace: "맑은 고딕", bold: true, color: green, margin: 0
});
const mechSports = ["웨이트/저항 트레이닝", "필라테스/요가", "배드민턴"];
mechSports.forEach((s, i) => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.55, y: 1.82 + i * 0.28, w: 2.7, h: 0.24,
    fill: { color: "FFFFFF" }, line: { color: green, width: 0.75 }, rectRadius: 0.04
  });
  slide.addText(s, {
    x: 3.55, y: 1.82 + i * 0.28, w: 2.7, h: 0.24,
    fontSize: 9, fontFace: "맑은 고딕", color: darkGray, align: "center", valign: "middle", margin: 0
  });
});

// 생리적 분석 group
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.4, y: 2.85, w: 3.0, h: 0.95,
  fill: { color: lightOrange }, line: { color: orange, width: 1 }, rectRadius: 0.06
});
slide.addText("생리적 분석", {
  x: 3.5, y: 2.85, w: 1.2, h: 0.3,
  fontSize: 8, fontFace: "맑은 고딕", bold: true, color: orange, margin: 0
});
const physSports = ["실내 사이클링", "보행/트레드밀"];
physSports.forEach((s, i) => {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 3.55, y: 3.17 + i * 0.28, w: 2.7, h: 0.24,
    fill: { color: "FFFFFF" }, line: { color: orange, width: 0.75 }, rectRadius: 0.04
  });
  slide.addText(s, {
    x: 3.55, y: 3.17 + i * 0.28, w: 2.7, h: 0.24,
    fontSize: 9, fontFace: "맑은 고딕", color: darkGray, align: "center", valign: "middle", margin: 0
  });
});

// === RIGHT COLUMN: AI 추론 결과 ===
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 7.0, y: 0.9, w: 2.7, h: 0.45,
  fill: { color: navy }, rectRadius: 0.08
});
slide.addText("AI 추론 결과 (이상 징후)", {
  x: 7.0, y: 0.9, w: 2.7, h: 0.45,
  fontSize: 12, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// 역학적 추론 결과
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.9, y: 1.5, w: 2.9, h: 1.2,
  fill: { color: lightGreen }, line: { color: green, width: 1 }, rectRadius: 0.06
});
slide.addText([
  { text: "근골격계 불안정성", options: { bold: true, fontSize: 10, color: green, breakLine: true } },
  { text: "좌우 근력 불균형", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "관절 가동범위(ROM) 이상", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "자세 정렬 이상·균형 능력 저하", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "상지 궤적 비대칭", options: { fontSize: 9, color: darkGray } }
], {
  x: 7.05, y: 1.55, w: 2.6, h: 1.1,
  fontFace: "맑은 고딕", valign: "top", margin: 0
});

// 생리적 추론 결과
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 6.9, y: 2.85, w: 2.9, h: 0.95,
  fill: { color: lightOrange }, line: { color: orange, width: 1 }, rectRadius: 0.06
});
slide.addText([
  { text: "대사적 활성도 추론", options: { bold: true, fontSize: 10, color: orange, breakLine: true } },
  { text: "심박 회복률(HRR) 이상", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "보행 효율·낙상 위험도", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "심폐 지구력·VO2 추론", options: { fontSize: 9, color: darkGray } }
], {
  x: 7.05, y: 2.9, w: 2.6, h: 0.85,
  fontFace: "맑은 고딕", valign: "top", margin: 0
});

// === ARROWS (left → middle) ===
const arrowOpts1 = { line: { color: blue, width: 2, endArrowType: "triangle" } };
slide.addShape(pres.shapes.LINE, { x: 2.9, y: 2.0, w: 0.5, h: 0, ...arrowOpts1 });
const arrowOpts1b = { line: { color: blue, width: 2, endArrowType: "triangle" } };
slide.addShape(pres.shapes.LINE, { x: 2.9, y: 3.2, w: 0.5, h: 0, ...arrowOpts1b });

// === ARROWS (middle → right) ===
const arrowOpts2 = { line: { color: green, width: 2, endArrowType: "triangle" } };
slide.addShape(pres.shapes.LINE, { x: 6.4, y: 2.1, w: 0.5, h: 0, ...arrowOpts2 });
const arrowOpts2b = { line: { color: orange, width: 2, endArrowType: "triangle" } };
slide.addShape(pres.shapes.LINE, { x: 6.4, y: 3.3, w: 0.5, h: 0, ...arrowOpts2b });

// === BOTTOM: 복합 분석 ===
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 2.5, y: 4.15, w: 5.3, h: 1.2,
  fill: { color: lightPurple }, line: { color: purple, width: 1.5 }, rectRadius: 0.08
});
slide.addText("복합 분석 (종목 조합 → 통합 신체 기능 평가)", {
  x: 2.5, y: 4.15, w: 5.3, h: 0.35,
  fontSize: 11, fontFace: "맑은 고딕", bold: true, color: purple, align: "center", valign: "middle", margin: 0
});

const combos = [
  { left: "웨이트 + 사이클링", right: "근력 대비 심폐 기능 균형" },
  { left: "보행 + 배드민턴", right: "하지/상지 통합 기능 평가" },
];
combos.forEach((c, i) => {
  const y = 4.55 + i * 0.35;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 2.7, y: y, w: 2.0, h: 0.28,
    fill: { color: "FFFFFF" }, line: { color: purple, width: 0.75 }, rectRadius: 0.04
  });
  slide.addText(c.left, {
    x: 2.7, y: y, w: 2.0, h: 0.28,
    fontSize: 9, fontFace: "맑은 고딕", color: darkGray, align: "center", valign: "middle", margin: 0
  });
  // arrow
  const arrowC = { line: { color: purple, width: 1.5, endArrowType: "triangle" } };
  slide.addShape(pres.shapes.LINE, { x: 4.75, y: y + 0.14, w: 0.5, h: 0, ...arrowC });
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.3, y: y, w: 2.3, h: 0.28,
    fill: { color: "FFFFFF" }, line: { color: purple, width: 0.75 }, rectRadius: 0.04
  });
  slide.addText(c.right, {
    x: 5.3, y: y, w: 2.3, h: 0.28,
    fontSize: 9, fontFace: "맑은 고딕", color: darkGray, align: "center", valign: "middle", margin: 0
  });
});

// Vertical arrows to 복합 분석 box
const vArrow1 = { line: { color: purple, width: 1.5, endArrowType: "triangle", dashType: "dash" } };
slide.addShape(pres.shapes.LINE, { x: 4.9, y: 3.85, w: 0, h: 0.3, ...vArrow1 });

// Footer note
slide.addText("※ 추론 정확도 목표: 1차년도 80% → 2차년도 85% → 3차년도 90% (ICC 기준, Doctor-Verified Set 검증)", {
  x: 0.5, y: 5.15, w: 9, h: 0.3,
  fontSize: 9, fontFace: "맑은 고딕", color: gray, margin: 0
});

pres.writeFile({ fileName: "C:/Users/박동혁/Desktop/Claude 연구용/연구비/2026/문체부 과제/PSR_AI추론_매핑도식.pptx" })
  .then(() => console.log("Done"))
  .catch(e => console.error(e));

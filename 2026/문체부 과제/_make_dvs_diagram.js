const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Doctor-Verified Set 구축 전략";

const slide = pres.addSlide();
slide.background = { color: "FFFFFF" };

const navy = "1E3A5F";
const blue = "4472C4";
const lightBlue = "D6E4F0";
const green = "2E7D32";
const lightGreen = "E8F5E9";
const orange = "E65100";
const lightOrange = "FFF3E0";
const red = "C62828";
const lightRed = "FFEBEE";
const gray = "546E7A";
const darkGray = "263238";

// Title
slide.addText("Doctor-Verified Set 구축 전략", {
  x: 0.5, y: 0.2, w: 9, h: 0.5,
  fontSize: 20, fontFace: "맑은 고딕", bold: true, color: navy, margin: 0
});

// === BOX 1: Ground Truth (연구소 정밀 장비) ===
const box1Y = 1.0;
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 2.0, y: box1Y, w: 6.0, h: 1.1,
  fill: { color: lightBlue }, line: { color: blue, width: 2 }, rectRadius: 0.1
});
slide.addText([
  { text: "연구소 정밀 장비 측정 결과 = 정답 (Ground Truth)", options: { bold: true, fontSize: 13, color: navy, breakLine: true } },
  { text: "Motion Capture  |  Isokinetic Dynamometer  |  Metabolic Test  |  EMG  |  Force Plate", options: { fontSize: 10, color: gray } }
], {
  x: 2.2, y: box1Y + 0.1, w: 5.6, h: 0.9,
  fontFace: "맑은 고딕", align: "center", valign: "middle", margin: 0
});

// Icon label left
slide.addShape(pres.shapes.OVAL, {
  x: 0.5, y: box1Y + 0.15, w: 0.8, h: 0.8,
  fill: { color: blue }
});
slide.addText("Gold\nStandard", {
  x: 0.5, y: box1Y + 0.15, w: 0.8, h: 0.8,
  fontSize: 9, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// === 동시 측정 화살표 ===
const arrowY = 2.25;
slide.addShape(pres.shapes.LINE, {
  x: 5.0, y: arrowY, w: 0, h: 0.6,
  line: { color: red, width: 3, endArrowType: "triangle" }
});
// 동시 측정 label
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 3.5, y: arrowY + 0.1, w: 1.3, h: 0.4,
  fill: { color: lightRed }, line: { color: red, width: 1 }, rectRadius: 0.06
});
slide.addText("동시 측정", {
  x: 3.5, y: arrowY + 0.1, w: 1.3, h: 0.4,
  fontSize: 11, fontFace: "맑은 고딕", bold: true, color: red, align: "center", valign: "middle", margin: 0
});

// === BOX 2: 웨어러블 수집 데이터 ===
const box2Y = 3.0;
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 2.0, y: box2Y, w: 6.0, h: 1.1,
  fill: { color: lightGreen }, line: { color: green, width: 2 }, rectRadius: 0.1
});
slide.addText([
  { text: "웨어러블 수집 데이터 = AI 입력", options: { bold: true, fontSize: 13, color: green, breakLine: true } },
  { text: "Smart Watch (가속도/자이로/심박)  |  S-Patch (HRR/HRV)  |  Bio Armour (ROM/근력)", options: { fontSize: 10, color: gray } }
], {
  x: 2.2, y: box2Y + 0.1, w: 5.6, h: 0.9,
  fontFace: "맑은 고딕", align: "center", valign: "middle", margin: 0
});

// Icon label left
slide.addShape(pres.shapes.OVAL, {
  x: 0.5, y: box2Y + 0.15, w: 0.8, h: 0.8,
  fill: { color: green }
});
slide.addText("AI\nInput", {
  x: 0.5, y: box2Y + 0.15, w: 0.8, h: 0.8,
  fontSize: 9, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// === 하단 화살표 ===
const arrow2Y = 4.25;
slide.addShape(pres.shapes.LINE, {
  x: 5.0, y: arrow2Y, w: 0, h: 0.5,
  line: { color: orange, width: 3, endArrowType: "triangle" }
});

// === BOX 3: 최종 목표 ===
const box3Y = 4.85;
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 2.0, y: box3Y, w: 6.0, h: 0.65,
  fill: { color: lightOrange }, line: { color: orange, width: 2.5 }, rectRadius: 0.1
});
slide.addText([
  { text: "AI가 웨어러블 데이터만으로 정밀 장비 수준의 추론을 달성하는 것이 목표", options: { bold: true, fontSize: 12, color: orange, breakLine: true } },
  { text: "(추론 정확도 90% 이상, ICC 기준, Bland-Altman 분석 병행)", options: { fontSize: 10, color: gray } }
], {
  x: 2.2, y: box3Y, w: 5.6, h: 0.65,
  fontFace: "맑은 고딕", align: "center", valign: "middle", margin: 0
});

// Icon label left - target
slide.addShape(pres.shapes.OVAL, {
  x: 0.5, y: box3Y, w: 0.8, h: 0.65,
  fill: { color: orange }
});
slide.addText("목표\n90%+", {
  x: 0.5, y: box3Y, w: 0.8, h: 0.65,
  fontSize: 9, fontFace: "맑은 고딕", bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0
});

// Right side: 검증 방법 callout
slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
  x: 8.3, y: 1.8, w: 1.5, h: 2.6,
  fill: { color: "F5F5F5" }, line: { color: gray, width: 1 }, rectRadius: 0.08
});
slide.addText([
  { text: "검증 방법", options: { bold: true, fontSize: 10, color: navy, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "ICC 산출", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "(급내상관계수)", options: { fontSize: 8, color: gray, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "Bland-Altman", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "(일치도 분석)", options: { fontSize: 8, color: gray, breakLine: true } },
  { text: "", options: { fontSize: 6, breakLine: true } },
  { text: "RMSE", options: { fontSize: 9, color: darkGray, breakLine: true } },
  { text: "(오차 정량화)", options: { fontSize: 8, color: gray } }
], {
  x: 8.4, y: 1.9, w: 1.3, h: 2.4,
  fontFace: "맑은 고딕", align: "center", valign: "top", margin: 0
});

pres.writeFile({ fileName: "C:/Users/박동혁/Desktop/Claude 연구용/연구비/2026/문체부 과제/DVS_구축전략_도식.pptx" })
  .then(() => console.log("Done"))
  .catch(e => console.error(e));

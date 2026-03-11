const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();

  // 가-2. 평가방법 및 평가환경
  const ws1 = wb.addWorksheet('가-2 평가방법');
  const h1 = ws1.addRow(['주요 성능지표', '세부 평가방법 및 평가환경']);
  h1.height = 28;
  h1.eachCell(cell => {
    cell.font = { name: '맑은 고딕', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  });

  const evalData = [
    ['PSR 과제 상호운용 API', 'API 호출 로그 기반 성공률 자동 산출. 총괄협의체 연계 과제 간 상호운용 테스트 환경에서 평가. 호출 성공률 = (성공 응답 수 / 전체 호출 수) × 100'],
    ['운동 처방 일치도 (ICC)', '연세대 FRICSS 정밀 장비(등속성 근력계, Motion Capture, 대사 분석기)로 구축한 Gold Standard와 웨어러블 AI 추론 결과 간 ICC 산출. Bland-Altman 분석 병행. 평가 대상: 5종 실증 종목별 최소 30명 이상'],
    ['데이터 가용성', 'CHAGEUN Platform 서버 로그 기반 서비스 호출 성공률 측정. 실시간 데이터 수집·전송·저장 파이프라인의 End-to-End 가용성 평가. 평가 기간: 연속 30일 이상 운영 환경'],
    ['사용성평가 만족도', '실제 사용자(운동참가자, 코치, 전문가) 대상 설문조사 실시. 총괄협의체 제시 표준화 만족도 척도 사용. 평가 대상: 소비자 50명 이상, 전문가 10명 이상'],
    ['AI 머신러닝 분류성능평가', 'ISO/IEC TS 4213 기준에 따라 공인시험·인증기관에서 평가 수행. 동작추적 정확도, 재현성, 환경조건 변화 대응력 등 포함'],
    ['AI 데이터 품질평가', 'ISO/IEC 5259-2 기준에 따라 공인시험·인증기관에서 평가 수행. 데이터 완전성, 정확성, 일관성, 적시성 등 품질 모델 기반 측정'],
  ];

  evalData.forEach(rowData => {
    const row = ws1.addRow(rowData);
    row.height = 50;
    row.eachCell((cell, colNum) => {
      cell.font = { name: '맑은 고딕', size: 10 };
      cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });

  ws1.getColumn(1).width = 28;
  ws1.getColumn(2).width = 72;

  // 나. 성과물 목표
  const ws2 = wb.addWorksheet('나 성과물목표');
  const h2 = ws2.addRow(['단계', '성과물 목표']);
  h2.height = 28;
  h2.eachCell(cell => {
    cell.font = { name: '맑은 고딕', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
  });

  const stage1Items = [
    '○ 의료/스포츠 전문가 검증 신체 기능 추론 벤치마크 데이터셋 (Doctor-Verified Set)',
    '○ 개인별 맞춤형 운동 가이드 프로그램 프로토타입',
    '○ AI 기반 건강 상태 평가 모델 (역학적·생리적 추론 정확도 85% 이상)',
    '○ PSR 과제 상호운용 API 4종 (호출 성공률 85% 이상)',
    '○ 스포츠 지식 온톨로지 및 LLM 에이전트 연동 API 구조',
  ];

  const stage2Items = [
    '○ 신체 기능 분석 및 운동 가이드 AI 엔진 소프트웨어 (추론 정확도 90% 이상)',
    '○ 사용자 피드백 기반 지속학습형 AI 추천 시스템',
    '○ PSR 과제 상호운용 API 6종 완비 (호출 성공률 95% 이상)',
    '○ AI 머신러닝 분류 성능평가 1건 (ISO/IEC TS 4213)',
    '○ AI 데이터 품질평가 1건 (ISO/IEC 5259-2)',
  ];

  const postItems = [
    '○ 기술이전 1건 이상',
    '○ 사업화 (B2B/B2C 서비스 모델)',
  ];

  const stageData = [
    ['1단계', stage1Items.join('\n')],
    ['2단계', stage2Items.join('\n')],
    ['과제 종료 후\n(2년 이내)', postItems.join('\n')],
  ];

  stageData.forEach(rowData => {
    const row = ws2.addRow(rowData);
    row.height = rowData[0] === '1단계' ? 90 : rowData[0] === '2단계' ? 90 : 45;
    row.eachCell((cell, colNum) => {
      cell.font = { name: '맑은 고딕', size: 10, bold: colNum === 1 };
      cell.alignment = { horizontal: colNum === 1 ? 'center' : 'left', vertical: 'middle', wrapText: true };
      cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
    });
  });

  ws2.getColumn(1).width = 16;
  ws2.getColumn(2).width = 72;

  await wb.xlsx.writeFile('성능지표_가2_나.xlsx');
  console.log('Done: 성능지표_가2_나.xlsx');
}

main();

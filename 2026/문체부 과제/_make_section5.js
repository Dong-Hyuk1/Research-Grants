const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  const thin = {style:'thin'};
  const border = {top:thin,left:thin,bottom:thin,right:thin};
  const hFont = {name:'맑은 고딕',bold:true,size:10,color:{argb:'FFFFFFFF'}};
  const hFill = {type:'pattern',pattern:'solid',fgColor:{argb:'FF4472C4'}};
  const dFont = {name:'맑은 고딕',size:10};
  const bFont = {name:'맑은 고딕',size:10,bold:true};
  const cAlign = {horizontal:'center',vertical:'middle',wrapText:true};
  const lAlign = {horizontal:'left',vertical:'middle',wrapText:true};

  const ws = wb.addWorksheet('5-3 사업화전략');

  const data = [
    ['사업화 모델(BM)',
     '① CHAGEUN 구독형 서비스: B2C 개인 사용자 대상 월 구독형 AI 운동 가이드 서비스 (모바일 앱)\n② CHAGEUN Enterprise: B2B 체육시설·피트니스 센터 대상 SaaS 솔루션 (센서 + AI 분석 패키지)\n③ CHAGEUN Medical: 의료기관·재활병원 대상 운동 처방 보조 도구 (의료기기급 서비스)'],
    ['사업화 추진 주체',
     'HAII Corp. (주관기관, SW 플랫폼 개발·운영)\nWellysis (공동, 웨어러블 센서 HW 제조·판매)'],
    ['시장분석',
     '① 글로벌 웨어러블 헬스케어 시장: 2028년 약 $186B 규모 (CAGR 14.6%)\n② 국내 디지털 헬스케어 시장: 2027년 약 4.5조 원 규모\n③ 경쟁 현황: Apple Watch, Garmin 등 글로벌 웨어러블은 단일 센서 기반 모니터링에 한정, 역학적·생리적 통합 추론 AI 보유 경쟁사 부재\n④ 시장 진입 장벽: Doctor-Verified Set, 다기관 임상 실증 데이터, ISO 국제표준 인증이 핵심 차별화 요소'],
    ['사업화 전략 및\n계획',
     '① 1단계 (과제 수행 중): 다기관 실증 데이터 확보 및 제품 신뢰성 근거 구축, PSR 생태계 내 API 연동을 통한 초기 사용자 기반 확보\n② 2단계 (과제 종료~1년): B2B 파일럿(체육시설 10개소, 의료기관 5개소) 시장 검증 및 BM 최적화, 연세대 기술이전 완료\n③ 3단계 (과제 종료 후 1~2년): B2C 앱 정식 출시 및 B2B 확장(체육시설 500개소, 의료기관 50개소), 해외 시장 진출 준비\n④ 수익 모델: B2C 월 구독료(9,900~19,900원), B2B SaaS 월 이용료(50~200만 원/시설), Medical 라이선스(기관당 연 1,000만 원)\n⑤ 매출 목표: 사업화 후 5년 내 누적 매출 100억 원'],
  ];

  data.forEach(rd => {
    const row = ws.addRow(rd);
    row.height = rd[0].includes('시장분석') || rd[0].includes('사업화 전략') ? 100 : 65;
    row.eachCell((c,i) => {
      c.font = i===1 ? bFont : dFont;
      c.alignment = i===1 ? cAlign : lAlign;
      c.border = border;
      if (i===1) c.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD9E2F3'}};
    });
  });

  ws.getColumn(1).width = 18;
  ws.getColumn(2).width = 80;

  await wb.xlsx.writeFile('사업화전략_5-3.xlsx');
  console.log('Done');
}
main();

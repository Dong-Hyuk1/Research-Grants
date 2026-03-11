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

  function addHeader(ws, headers, h=30) {
    const row = ws.addRow(headers);
    row.height = h;
    row.eachCell(c => { c.font=hFont; c.fill=hFill; c.alignment=cAlign; c.border=border; });
  }
  function addData(ws, data, opts={}) {
    data.forEach(rd => {
      const row = ws.addRow(rd);
      row.height = opts.h || 30;
      row.eachCell((c,i) => {
        c.font = (opts.boldCols && opts.boldCols.includes(i)) ? bFont : dFont;
        c.alignment = (opts.leftCols && opts.leftCols.includes(i)) ? lAlign : cAlign;
        c.border = border;
      });
    });
  }

  // ===== 3-2. 추진체계 =====
  const ws1 = wb.addWorksheet('3-2 추진체계');
  addHeader(ws1, ['기관명', '담당 연구개발내용', '연구개발 역량의 우수성']);

  const orgData = [
    ['주관연구개발기관\nHAII Corp.\n(정훈엽 소장)',
     '○ PSR 기반 맞춤형 운동 가이드 생성 시스템 SW 개발 및 사업화\n○ AI 추론 엔진 통합 개발, CHAGEUN Platform 구축·운영\n○ PSR 과제 상호운용 API 개발\n○ 총괄협의체 데이터 표준화 인터페이스 구현',
     '○ 디지털 바이오마커 기반 근감소증 디지털 치료기기 개발 과제(42억 원) 수행 중\n○ 9종 디지털 바이오마커 개발 완료\n○ Google ML Kit Pose Detection 기반 ML 파이프라인 구축\n○ CHAGEUN Platform 보유'],
    ['공동연구개발기관\n연세대학교 산학협력단\n(전용관 교수)',
     '○ PSR 데이터 기반 신체 기능 예측 모델 설계\n○ 맞춤형 운동 가이드 알고리즘 설계\n○ FRICSS 정밀 장비 활용 Gold Standard 데이터셋 구축\n○ Doctor-Verified Set 확보\n○ 국제표준 기반 성능·품질 평가 수행',
     '○ FRICSS 연구소 운영 (등속성 근력계, 3D Motion Capture, EMG, 대사 분석기 등 완비)\n○ Nature Medicine 포함 200편+ 논문\n○ 15개 질환 운동프로그램 개발 경험\n○ 5개 운동부(150명) 데이터 접근 가능'],
    ['공동연구개발기관\nWellysis\n(김종우 대표)',
     '○ Smart Watch + S-Patch + Bio Armour 3종 웨어러블 센서 기반 PSR 데이터 수집 시스템 구축·유지보수\n○ 센서 간 상호호환 데이터 통합 파이프라인 구축\n○ 실증 현장 기술 지원',
     '○ S-Patch FDA 승인 획득\n○ Bio Armour(비침습적 ROM·근력 측정) 자체 개발\n○ 멀티모달 센서 상호호환 시스템 기 구축\n○ 웨어러블 헬스케어 기기 사업화 역량'],
    ['필수 공동연구개발기관\n신촌 세브란스병원\n(3차 의료기관)',
     '○ 일반 성인 대상 근골격계 신체 증상 기전 제공\n○ 웨이트 트레이닝·실내 사이클링 종목 실증\n○ AI 추론 결과 의료 전문가 일치도 검증\n○ Doctor-Verified Set 구축 참여',
     '○ 국내 최상위 3차 의료기관\n○ 다기관 임상시험 수행 인프라 완비\n○ 근골격계 질환 전문 진료 역량\n○ 디지털 치료기기 과제 임상 참여 경험'],
    ['필수 공동연구개발기관\n경희대학교 의료원\n(원장원 교수, 3차 의료기관)',
     '○ 노인·갱년기 여성 대상 필라테스/요가·보행 실증\n○ Bio Armour 기반 대사 효율성 지표 도출\n○ 고령층 특화 운동 가이드 효과성 검증\n○ PSR 기반 맞춤형 운동 솔루션 자문',
     '○ 노인·갱년기 여성 운동 처방 연구 경험\n○ 고령층 대사 효율성 평가 역량\n○ 디지털 치료기기 과제 임상 참여 경험'],
    ['필수 공동연구개발기관\n충남대학교 병원\n(문창원 교수, 3차 의료기관)',
     '○ 신장재활 환자 대상 보행/트레드밀·실내 사이클링 실증\n○ 재활 기준 근골격계 신체 증상 기전 제공\n○ 만성질환 대상자 운동 안전성·효과성 검증',
     '○ 신장재활 환자 재활 운동 프로그램 연구 경험\n○ 만성질환(CKD) 재활 전문 진료 역량\n○ 디지털 치료기기 과제 임상 참여 경험'],
  ];

  orgData.forEach(rd => {
    const row = ws1.addRow(rd);
    row.height = 85;
    row.eachCell((c,i) => {
      c.font = i===1 ? bFont : dFont;
      c.alignment = lAlign;
      c.border = border;
    });
  });
  ws1.getColumn(1).width = 24;
  ws1.getColumn(2).width = 45;
  ws1.getColumn(3).width = 45;

  // ===== 3-3 추진일정 (3 sheets) =====
  function makeGantt(ws, title, months, data) {
    const headers = ['연구개발내용', ...months.map(String)];
    addHeader(ws, headers, 25);
    const fillGray = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD9E2F3'}};
    data.forEach(rd => {
      const row = ws.addRow([rd[0], ...rd.slice(1)]);
      row.height = 22;
      row.eachCell((c,i) => {
        c.font = dFont;
        c.alignment = i===1 ? {...lAlign,wrapText:false} : cAlign;
        c.border = border;
        if (i > 1 && c.value === '■') {
          c.fill = fillGray;
          c.font = {name:'맑은 고딕',size:10,bold:true,color:{argb:'FF2F5496'}};
        }
      });
    });
    ws.getColumn(1).width = 42;
    for (let i=2; i<=months.length+1; i++) ws.getColumn(i).width = 5;
  }

  // 1차년도
  const ws2 = wb.addWorksheet('1단계 1차년도');
  makeGantt(ws2, '1차년도', [4,5,6,7,8,9,10,11,12], [
    ['실증 인프라 구축 (IRB, 센서 셋업, 프로토콜)','■','■','■','','','','','',''],
    ['PSR 데이터 수집 프로토콜 확정 및 파일럿 테스트','','■','■','','','','','',''],
    ['FRICSS 정밀장비 + 웨어러블 동시측정 (Ground Truth)','','','■','■','■','■','','',''],
    ['Doctor-Verified Set 구축 (전문가 검증)','','','','■','■','■','','',''],
    ['역학적 추론 모델 프로토타입 개발','','','','','■','■','■','',''],
    ['생리적 추론 모델 프로토타입 개발','','','','','■','■','■','',''],
    ['운동 가이드 프로토타입 개발','','','','','','■','■','■',''],
    ['PSR 상호운용 API 2종 개발','','','','','','','■','■','■'],
    ['모델 검증 (추론 정확도 80%) 및 성과보고','','','','','','','','■','■'],
  ]);

  // 2차년도
  const ws3 = wb.addWorksheet('1단계 2차년도');
  makeGantt(ws3, '2차년도', [1,2,3,4,5,6,7,8,9,10,11,12], [
    ['다기관 실증 데이터 수집 (세브란스/경희대/충남대)','■','■','■','■','■','■','■','■','','','',''],
    ['운동부 보조 코호트 검증 (Known Asymmetry Group)','■','■','■','■','','','','','','','',''],
    ['AI 추론 모델 고도화 (정확도 85%)','','','■','■','■','■','■','','','','',''],
    ['맞춤형 운동 가이드 프로그램 개발','','','','■','■','■','■','■','','','',''],
    ['스포츠 지식 온톨로지 구축','','','','','■','■','■','■','■','','',''],
    ['LLM 에이전트 연동 API 설계','','','','','','■','■','■','■','','',''],
    ['적응형 운동 강도 조정 알고리즘 개발','','','','','','','■','■','■','■','',''],
    ['PSR 상호운용 API 4종 확장','','','','','','','','■','■','■','',''],
    ['사용성 평가 (만족도 85%) 및 성과보고','','','','','','','','','','■','■','■'],
  ]);

  // 3차년도
  const ws4 = wb.addWorksheet('2단계 3차년도');
  makeGantt(ws4, '3차년도', [1,2,3,4,5,6,7,8,9,10,11,12], [
    ['통합 AI 추론 엔진 고도화 (정확도 90%)','■','■','■','■','■','■','','','','','',''],
    ['다기관 리빙랩 실증','■','■','■','■','■','■','■','■','','','',''],
    ['지속학습형 AI 추천 시스템 완성','','','■','■','■','■','■','■','','','',''],
    ['PSR 상호운용 API 6종 완비 (성공률 95%)','','','','','■','■','■','■','■','','',''],
    ['AI 분류 성능평가 (ISO/IEC TS 4213)','','','','','','','■','■','■','■','',''],
    ['AI 데이터 품질평가 (ISO/IEC 5259-2)','','','','','','','■','■','■','■','',''],
    ['사용성 평가 (만족도 90%)','','','','','','','','■','■','■','',''],
    ['기술이전 준비 및 최종 성과보고','','','','','','','','','','■','■','■'],
  ]);

  // ===== 웨어러블 센서 구성표 =====
  const ws5 = wb.addWorksheet('웨어러블 센서 구성');
  addHeader(ws5, ['수집 장비', '수집 데이터', '분석 유형']);
  addData(ws5, [
    ['Smart Watch', '가속도/자이로(상지 움직임), 보행 패턴, 심박, BIA 근육량', '역학적 + 생리적'],
    ['S-Patch (FDA 승인)', '심박 회복률(HRR), HRV, 파워 출력 지속성', '생리적 분석'],
    ['Bio Armour', '관절 가동범위(ROM), 근력(Arm Curl, STS)', '역학적 분석'],
    ['Mobile Software', '사용자 인터페이스, 데이터 전송, Pose Detection', '역학적 분석'],
  ], {h:30, boldCols:[1], leftCols:[1,2]});
  ws5.getColumn(1).width = 22;
  ws5.getColumn(2).width = 50;
  ws5.getColumn(3).width = 18;

  // ===== 실증 종목 5종 =====
  const ws6 = wb.addWorksheet('실증 종목 5종');
  addHeader(ws6, ['순위', '종목', '분석 유형', '주요 웨어러블', '연구소 검증 장비\n(Ground Truth)', '선정 근거']);
  const sportData = [
    ['1', '웨이트/저항\n트레이닝', '역학적\n+ 생리적', 'Bio Armour(ROM, 근력)\nWatch(Kinetic)', 'Isokinetic Dynamometer\nEMG, Motion Capture', '등속성 근력계로 좌우 불균형 Gold Standard 제공.\nBio Armour 추론 정확도를 Isokinetic으로 검증.'],
    ['2', '필라테스/요가', '역학적', 'Bio Armour(ROM)\nWatch(균형)', 'Motion Capture\nGoniometer\nBalance Force Plate', 'Motion Capture로 관절 가동범위 Ground Truth 확보.\n갱년기 여성/노인 대상(경희대 실증)과 직결.'],
    ['3', '실내 사이클링', '생리적', 'S-Patch(HRR, HRV)\nWatch(심박)', 'Cycle Ergometer\nStress Test(VO2max)\nMetabolic Test', 'Metabolic Test로 대사적 활성도 Gold Standard 제공.\nS-Patch 추론을 VO2/대사율로 검증.'],
    ['4', '보행/트레드밀\n워킹', '역학적\n+ 생리적', 'Watch(보행, 균형)\nBio Armour(하지ROM)', 'Force Plate\nFoot Switch\nMotion Capture', 'Force Plate + Foot Switch로 보행 기능 Gold Standard 확보.\n노인/재활(충남대) 실증 필수 종목.'],
    ['5', '배드민턴', '역학적\n+ 생리적', 'Watch(상지 Kinetic)\nS-Patch(심박)', 'Motion Capture\nEMG(상지)\nForce Plate', 'Motion Capture로 상지 운동 궤적 Gold Standard 확보.\n좌우 불균형(주사용 팔) 뚜렷.'],
  ];
  sportData.forEach(rd => {
    const row = ws6.addRow(rd);
    row.height = 50;
    row.eachCell((c,i) => {
      c.font = i===2 ? bFont : dFont;
      c.alignment = i<=3 ? cAlign : lAlign;
      c.border = border;
    });
  });
  ws6.getColumn(1).width = 6;
  ws6.getColumn(2).width = 14;
  ws6.getColumn(3).width = 10;
  ws6.getColumn(4).width = 22;
  ws6.getColumn(5).width = 24;
  ws6.getColumn(6).width = 40;

  // ===== Ground Truth 검증 체계 =====
  const ws7 = wb.addWorksheet('Ground Truth 검증체계');
  addHeader(ws7, ['검증 영역', '정밀 장비 (Gold Standard)', '웨어러블 (현장 수집)', '비교 방법']);
  addData(ws7, [
    ['근력/좌우 균형', 'Isokinetic Dynamometer', 'Bio Armour (Arm Curl, STS)', 'ICC, Bland-Altman'],
    ['근활성도', 'EMG System', 'Bio Armour + Watch (Kinetic)', '상관분석, 패턴 매칭'],
    ['동작 궤적', 'Motion Capture System (3D)', 'Watch (가속도/자이로) + Pose Detection', 'RMSE, 궤적 유사도'],
    ['관절 가동범위', 'Goniometer + Motion Capture', 'Bio Armour (ROM)', 'ICC, LOA'],
    ['균형/안정성', 'Balance Force Plate', 'Watch (가속도) + Pose Detection (OLS)', 'COP 분석, 상관분석'],
    ['심폐/대사 기능', 'Stress Test (VO2max) + Metabolic Test', 'S-Patch (HRR, HRV) + Watch (심박)', '회귀분석, ICC'],
    ['보행 기능', 'Force Plate + Foot Switch', 'Watch (보행패턴) + Pose Detection (TUG)', '보행 파라미터 비교'],
  ], {h:30, boldCols:[1], leftCols:[1,2,3,4]});
  ws7.getColumn(1).width = 16;
  ws7.getColumn(2).width = 32;
  ws7.getColumn(3).width = 36;
  ws7.getColumn(4).width = 22;

  await wb.xlsx.writeFile('섹션3_추진전략_표.xlsx');
  console.log('Done: 섹션3_추진전략_표.xlsx');
}

main();

const ExcelJS = require('exceljs');

async function main() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('성능지표');

  const headers = ['주요 성능지표', '단위', '1단계\n(1차년도)', '1단계\n(2차년도)', '2단계\n(3차년도)', '비교수준', '가중치', '평가방법'];

  const headerRow = ws.addRow(headers);
  headerRow.height = 32;
  headerRow.eachCell(cell => {
    cell.font = { name: '맑은 고딕', bold: true, size: 10, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
  });

  const data = [
    ['PSR 과제 상호운용 API\n(운영 API 수/호출성공률)', 'ea/%', '2/80%', '4/85%', '6/95%', '총괄협의체 기준', '20%', '자체평가'],
    ['운동 처방 일치도 (ICC)', 'ICC', '0.80', '0.85', '0.90', 'Doctor-Verified Set 기준', '25%', '공인시험·인증기관'],
    ['데이터 가용성', '%', '70', '85', '95', '실시간 서비스 응답성', '20%', '자체평가'],
    ['사용성평가\n(소비자/전문가) 만족도', '%', '-', '85', '90', '총괄협의체 제시 기준', '15%', '외부전문가 평가'],
    ['AI 머신러닝 분류성능평가', '건', '-', '-', '1', 'ISO/IEC TS 4213', '10%', '공인시험·인증기관'],
    ['AI 데이터 품질평가', '건', '-', '-', '1', 'ISO/IEC 5259-2', '10%', '공인시험·인증기관'],
  ];

  data.forEach(rowData => {
    const row = ws.addRow(rowData);
    row.height = 36;
    row.eachCell((cell, colNum) => {
      cell.font = { name: '맑은 고딕', size: 10 };
      cell.alignment = { horizontal: colNum === 1 ? 'left' : 'center', vertical: 'middle', wrapText: true };
      cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
    });
  });

  ws.getColumn(1).width = 32;
  ws.getColumn(2).width = 8;
  ws.getColumn(3).width = 14;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 14;
  ws.getColumn(6).width = 24;
  ws.getColumn(7).width = 10;
  ws.getColumn(8).width = 20;

  await wb.xlsx.writeFile('성능지표_가1.xlsx');
  console.log('Done: 성능지표_가1.xlsx');
}

main();

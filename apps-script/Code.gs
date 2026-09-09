/**
 * 2026 환경위기시계 실시간 설문 — Google Apps Script 백엔드
 *
 * 설치 방법
 * 1) 새 Google 스프레드시트를 만든다 (이름 예: "2026 환경위기시계 응답").
 * 2) 상단 메뉴 확장 프로그램(Extensions) > Apps Script 를 연다.
 * 3) 기본 생성된 코드를 지우고 이 파일 내용 전체를 붙여넣는다.
 * 4) 저장 후 상단 "배포(Deploy)" > "새 배포(New deployment)" 클릭.
 * 5) 유형 선택에서 "웹 앱(Web app)" 선택.
 *    - 실행 계정(Execute as): 나 (Me)
 *    - 액세스 권한(Who has access): 전체(Anyone)
 * 6) 배포 후 나오는 웹 앱 URL(.../exec 로 끝남)을 복사한다.
 * 7) 이 저장소의 config.js 파일의 SCRIPT_URL 값에 붙여넣고 저장한다.
 *
 * 참고: 최초 응답이 들어올 때 "Responses" 시트가 자동 생성된다.
 */

var SHEET_NAME = 'Responses';
var HEADERS = [
  'timestamp',
  'issue1_id', 'issue1_time',
  'issue2_id', 'issue2_time',
  'issue3_id', 'issue3_time',
  'score',
  'solutions',
  'other_text'
];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    var picks = body.picks || [];
    var p1 = picks[0] || {};
    var p2 = picks[1] || {};
    var p3 = picks[2] || {};

    sheet.appendRow([
      new Date(),
      p1.issueId || '', p1.time || '',
      p2.issueId || '', p2.time || '',
      p3.issueId || '', p3.time || '',
      body.score || '',
      (body.solutions || []).join('|'),
      body.other || ''
    ]);

    return jsonOut_({ ok: true });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  var sheet = getSheet_();
  var values = sheet.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    rows.push({
      ts: r[0] instanceof Date ? r[0].getTime() : r[0],
      picks: [
        { issueId: String(r[1] || ''), time: Number(r[2]) || 0 },
        { issueId: String(r[3] || ''), time: Number(r[4]) || 0 },
        { issueId: String(r[5] || ''), time: Number(r[6]) || 0 }
      ].filter(function (p) { return p.issueId; }),
      score: Number(r[7]) || 0,
      solutions: r[8] ? String(r[8]).split('|').filter(Boolean) : [],
      other: r[9] || ''
    });
  }

  return jsonOut_(rows);
}

/**
 * =====================================================================
 *  [젠파워 영어 웹사이트] 회원관리 & 로그인 연동 Google Apps Script
 *  스프레드시트: 젠파워영어회원관리
 *  스프레드시트 ID: 1FcfTZsKJ-C9_E4NuHOgJwc2DzOIZ563Hc7A1mqvMdhk
 *  스프레드시트 URL: https://docs.google.com/spreadsheets/d/1FcfTZsKJ-C9_E4NuHOgJwc2DzOIZ563Hc7A1mqvMdhk/edit
 * =====================================================================
 * 
 * [초간단 3단계 설정 가이드]
 * 1. 구글 스프레드시트(젠파워영어회원관리) 상단 메뉴:
 *    [확장 프로그램(Extensions)] -> [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 이 스크립트 전체를 복사하여 붙여넣기 후 저장 (Ctrl + S)
 * 3. 우측 상단 [배포(Deploy)] 버튼 클릭 -> [새 배포(New deployment)] 선택
 *    - 유형 선택(톱니바퀴): [웹 앱(Web app)]
 *    - 설명: 젠파워영어 회원 연동 API
 *    - 다음 사용자로 실행: 나 (Me)
 *    - 액세스 권한이 있는 사용자: 모든 사용자 (Anyone)  ★ 꼭 '모든 사용자'로 선택!
 * 4. [배포] 클릭 후 승인 절차를 마치면 나타나는 [웹 앱 URL] 복사!
 *    (예: https://script.google.com/macros/s/AKfycb.../exec)
 */

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // 동시 접속 락 대기

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.openById('1FcfTZsKJ-C9_E4NuHOgJwc2DzOIZ563Hc7A1mqvMdhk');
    }

    var params = {};
    if (e && e.postData && e.postData.contents) {
      try {
        params = JSON.parse(e.postData.contents);
      } catch (err) {
        params = e.parameter || {};
      }
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    var action = params.action || 'syncMember';
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

    // 1. 회원목록 시트 준비
    var memberSheet = ss.getSheetByName('회원목록');
    if (!memberSheet) {
      memberSheet = ss.insertSheet('회원목록', 0);
      var memberHeaders = [
        '가입일시', '회원유형', '이름', '이메일/아이디', '휴대폰번호', 
        '생년월일', '성별', '가입경로', '최근로그인일시', '로그인횟수', '상태'
      ];
      memberSheet.appendRow(memberHeaders);
      var headerRange = memberSheet.getRange(1, 1, 1, memberHeaders.length);
      headerRange.setBackground('#059669').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      memberSheet.setFrozenRows(1);
    }

    // 2. 접속로그 시트 준비
    var logSheet = ss.getSheetByName('접속로그');
    if (!logSheet) {
      logSheet = ss.insertSheet('접속로그', 1);
      var logHeaders = [
        '로그일시', '회원유형', '이름', '이메일/아이디', '휴대폰번호', 
        '로그인방식', '접속화면/서비스', '브라우저/환경'
      ];
      logSheet.appendRow(logHeaders);
      var logHeaderRange = logSheet.getRange(1, 1, 1, logHeaders.length);
      logHeaderRange.setBackground('#1e40af').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
      logSheet.setFrozenRows(1);
    }

    var userType = params.userType || '학생'; // 학생 or 학부모
    var name = params.name || '미입력';
    var email = params.email || params.id || '미입력';
    var phone = params.phone || '-';
    var birth = params.birth || '-';
    var gender = params.gender === 'M' ? '남성' : (params.gender === 'F' ? '여성' : (params.gender || '-'));
    var provider = params.provider || '이메일'; // 카카오, 네이버, Google, 이메일
    var service = params.service || '메인화면';
    var userAgent = params.userAgent || '-';

    // 회원목록에서 기존 회원 찾기 (이메일/아이디 기준 또는 휴대폰 번호 기준)
    var memberData = memberSheet.getDataRange().getValues();
    var rowIndex = -1;
    var loginCount = 1;

    for (var i = 1; i < memberData.length; i++) {
      var rowEmail = String(memberData[i][3]).trim();
      var rowPhone = String(memberData[i][4]).trim();
      if ((email !== '미입력' && rowEmail === email) || (phone !== '-' && rowPhone === phone)) {
        rowIndex = i + 1; // 1-based index
        loginCount = (parseInt(memberData[i][9], 10) || 1) + 1;
        break;
      }
    }

    if (rowIndex > 1) {
      // 기존 회원 정보 업데이트 (최근로그인, 로그인횟수 증가, 누락된 휴대폰/생년월일 보완)
      memberSheet.getRange(rowIndex, 9).setValue(nowStr); // 최근로그인일시
      memberSheet.getRange(rowIndex, 10).setValue(loginCount); // 로그인횟수
      if (phone !== '-' && (!memberData[rowIndex - 1][4] || memberData[rowIndex - 1][4] === '-')) {
        memberSheet.getRange(rowIndex, 5).setValue(phone);
      }
      if (birth !== '-' && (!memberData[rowIndex - 1][5] || memberData[rowIndex - 1][5] === '-')) {
        memberSheet.getRange(rowIndex, 6).setValue(birth);
      }
      if (gender !== '-' && (!memberData[rowIndex - 1][6] || memberData[rowIndex - 1][6] === '-')) {
        memberSheet.getRange(rowIndex, 7).setValue(gender);
      }
    } else {
      // 신규 회원 등록
      memberSheet.appendRow([
        nowStr, userType, name, email, phone, 
        birth, gender, provider, nowStr, 1, '정상'
      ]);
    }

    // 접속 로그 추가
    logSheet.appendRow([
      nowStr, userType, name, email, phone, 
      provider, service, userAgent
    ]);

    return createJsonResponse({
      success: true,
      message: '회원 정보 및 접속 로그가 구글 시트에 성공적으로 동기화되었습니다.',
      timestamp: nowStr,
      isNew: (rowIndex === -1)
    });

  } catch (err) {
    return createJsonResponse({
      success: false,
      error: err.toString()
    });
  } finally {
    lock.releaseLock();
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 최초 1회 실행하여 시트 헤더를 자동으로 생성하는 테스트 함수
function setupSheets() {
  handleRequest({
    parameter: {
      action: 'init',
      userType: '학생 (샘플)',
      name: '김도현',
      email: 'student@zenpower.edu',
      phone: '010-1234-5678',
      provider: '시스템초기화'
    }
  });
}

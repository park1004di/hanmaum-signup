/**
 * 2026 한마음 큰 잔치 — 신청폼 백엔드 (구글 시트 기록 + 접수 확인 메일 + 수정 지원)
 *
 * 이 파일 전체를 복사해서 사용하세요. 설정은 이미 채워져 있습니다.
 *  - SHEET_ID : '2026 한마음 큰 잔치 — 신청 접수 대장' 시트에 연결됨
 *  - SECRET   : 신청폼(index.html)·수정 페이지(edit.html)의 SECRET과 똑같이 유지
 *
 * 하는 일 (전부 무료):
 *  1) register : 신청 1건을 시트에 1행으로 기록 + 신청자에게 접수 확인 메일 발송
 *     (메일에 접수번호·비밀번호·수정 페이지 링크가 들어갑니다)
 *  2) lookup   : 이메일 + 비밀번호(해시)로 본인 신청 조회 — 수정 페이지가 사용
 *  3) update   : 본인 신청 내용 수정 (이메일·비밀번호 확인 후 해당 행만 갱신)
 */

var SHEET_ID = '1y5bnVssN8yRlyNU-rqP6gkJjKJhRYJlYoh-UTEtlNxM';
var SECRET   = 'hanmaum2026';
var SHEET_NAME = '신청접수';
var HEADERS = ['접수일시','접수번호','신청자','이메일','연락처','어른','어린이','합계','한인회비','참가자 명단','비밀번호(해시)','참가자데이터(JSON)','수정일시'];

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#101E38').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.secret !== SECRET) return out_({ ok: false, error: 'auth' });
    var sh = sheet_();

    if (d.action === 'register') {
      // 중복 기록 방지: 같은 접수번호가 이미 있으면 스킵
      var rows = sh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][1]) === String(d.receipt)) return out_({ ok: true, dup: true });
      }
      sh.appendRow([d.ts, d.receipt, d.rep, d.email, d.tel, d.adults, d.kids, d.total,
                    d.fee, d.roster, d.pwHash, d.members, '']);
      mail_(d); // 메일 발송이 실패해도 접수 기록은 유지됩니다
      return out_({ ok: true });
    }

    if (d.action === 'update') {
      var rows2 = sh.getDataRange().getValues();
      for (var j = 1; j < rows2.length; j++) {
        if (String(rows2[j][1]) === String(d.receipt) && String(rows2[j][10]) === String(d.pwHash)) {
          var r = j + 1;
          sh.getRange(r, 3).setValue(d.rep);                       // 신청자
          sh.getRange(r, 5).setValue(d.tel);                       // 연락처
          sh.getRange(r, 6, 1, 5).setValues([[d.adults, d.kids, d.total, d.fee, d.roster]]); // F~J
          sh.getRange(r, 12).setValue(d.members);                  // 참가자데이터(JSON)
          sh.getRange(r, 13).setValue(d.editTs);                   // 수정일시
          return out_({ ok: true });
        }
      }
      return out_({ ok: false, error: 'notfound' });
    }

    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    var p = e.parameter || {};
    if (p.action === 'lookup') {
      var sh = sheet_();
      var rows = sh.getDataRange().getValues();
      var wantEmail = String(p.email || '').trim().toLowerCase();
      var wantHash = String(p.pwhash || '');
      var list = [];
      for (var i = 1; i < rows.length; i++) {
        var rowEmail = String(rows[i][3]).trim().toLowerCase();
        if (rowEmail === wantEmail && String(rows[i][10]) === wantHash) {
          list.push({
            receipt: rows[i][1], ts: rows[i][0], rep: rows[i][2], tel: rows[i][4],
            adults: rows[i][5], kids: rows[i][6], total: rows[i][7], fee: rows[i][8],
            roster: rows[i][9], members: rows[i][11]
          });
        }
      }
      return out_({ ok: true, list: list }, p.callback);
    }
    return out_({ ok: false }, p.callback);
  } catch (err) {
    return out_({ ok: false, error: String(err) }, p.callback);
  }
}

/* 접수 확인 메일 — 신청자에게 발송 (비밀번호 + 수정 페이지 링크 포함) */
function mail_(d) {
  try {
    var link = d.editUrl || '(수정 페이지 링크는 곧 공지됩니다)';
    var html =
      '<div style="font-family:Malgun Gothic,Apple SD Gothic Neo,sans-serif;max-width:600px;line-height:1.7">'
      + '<h2 style="margin:0 0 6px;color:#101E38">접수가 완료되었습니다 ✓</h2>'
      + '<p style="margin:4px 0;color:#3A4A68">2026 한마음 큰 잔치 (2026. 11. 28. 토 · 양곤 실내 체육관)</p>'
      + '<table style="border-collapse:collapse;margin:14px 0">'
      + '<tr><td style="padding:4px 12px 4px 0;color:#7A869C">접수번호</td><td style="font-weight:900;color:#1F5AA8">' + d.receipt + '</td></tr>'
      + '<tr><td style="padding:4px 12px 4px 0;color:#7A869C">신청자</td><td>' + d.rep + '</td></tr>'
      + '<tr><td style="padding:4px 12px 4px 0;color:#7A869C">인원</td><td>총 ' + d.total + '명 (어른 ' + d.adults + ' · 어린이 ' + d.kids + ')</td></tr>'
      + '</table>'
      + '<pre style="background:#F6EED9;border-radius:8px;padding:12px;white-space:pre-wrap;font-size:13px">' + d.roster + '</pre>'
      + '<h3 style="margin:18px 0 6px;color:#101E38">내용을 수정하려면</h3>'
      + '<p style="margin:4px 0">아래 <b>수정 페이지</b>에서 <b>이 이메일</b>과 <b>비밀번호</b>를 입력해 로그인하세요.</p>'
      + '<p style="margin:4px 0">비밀번호: <b style="letter-spacing:1px">' + d.pw + '</b></p>'
      + '<p style="margin:10px 0"><a href="' + link + '" style="background:#1F5AA8;color:#fff;padding:11px 18px;border-radius:8px;text-decoration:none;font-weight:700">수정 페이지 열기</a></p>'
      + '<p style="margin:14px 0 0;font-size:13px;color:#7A869C">이 메일(비밀번호 포함)은 제3자에게 전달하지 마세요.<br>'
      + 'Registration confirmed — edit via the link above using your email and password.</p>'
      + '</div>';
    MailApp.sendEmail({ to: d.email, subject: '[한마음 큰 잔치] 접수 완료 — ' + d.receipt, htmlBody: html, name: '미얀마 한인회' });
  } catch (err) {
    // 메일 실패(주소 오류·일일 한도 초과 등)가 있어도 시트의 접수 기록은 그대로입니다
  }
}

function out_(j, cb) {
  var t = JSON.stringify(j);
  if (cb && /^[A-Za-z0-9_]+$/.test(cb)) t = cb + '(' + t + ')'; // JSONP — 수정 페이지에서 읽을 수 있게
  return ContentService.createTextOutput(t).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

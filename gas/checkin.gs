// ============================================================
// Seron LIFF – チェックイン機能
// ============================================================
// 【セットアップ手順】
//  1. cases シート T列 に checkin_code を追加（4桁文字列）
//  2. applications シート N列 に checkin_sent_at を追加
//  3. applications シート O列 に checkin_at を追加
//  4. LINE_CHANNEL_ACCESS_TOKEN を設定
//  5. dispatchCheckinButtons に5分おきの時間トリガーを設定
//  6. 既存の doGet / doPost に getApp / handleCheckin の呼び出しを追加
//  7. Apps Script エディタで「新しいデプロイ」を実行（必須）
//  8. プロジェクト設定でタイムゾーンが Asia/Tokyo になっているか確認
// ============================================================

// ---- 設定 ----
const LINE_CHANNEL_ACCESS_TOKEN = '★ここにチャネルアクセストークンを設定★';
const LIFF_BASE_URL = 'https://liff.line.me/2010386721-BjcO43xK';

// ---- applications シート 列インデックス (0始まり) ★実際の列と照合して修正★ ----
const COL_APP_ID           = 0;   // A: app_id
const COL_APP_CASE_ID      = 1;   // B: case_id
const COL_APP_USER_ID      = 2;   // C: user_id
const COL_APP_DISPLAY_NAME = 3;   // D: display_name
const COL_APP_DATE         = 4;   // E: requested_date
const COL_APP_TIME         = 5;   // F: requested_time
const COL_APP_STATUS       = 9;   // J: status ★要確認★
const COL_APP_NOSHOWFLAG   = 11;  // L: no_show
const COL_APP_CHECKIN_SENT = 13;  // N: checkin_sent_at (新規追加列)
const COL_APP_CHECKIN_AT   = 14;  // O: checkin_at      (新規追加列)

// ---- cases シート 列インデックス (0始まり) ★実際の列と照合して修正★ ----
const COL_CASE_ID           = 0;  // A: case_id
const COL_CASE_TITLE        = 1;  // B: title ★要確認★
const COL_CASE_AREA         = 2;  // C: area ★要確認★
const COL_CASE_CHECKIN_CODE = 19; // T: checkin_code (新規追加列)


// ============================================================
// ① ボタン自動送付（5分おきのトリガーで実行）
// ============================================================
function dispatchCheckinButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const appsSheet  = ss.getSheetByName('applications');
  const casesSheet = ss.getSheetByName('cases');

  // cases の checkin_code をマップ化
  const casesData = casesSheet.getDataRange().getValues();
  const caseMap = {};
  for (let i = 1; i < casesData.length; i++) {
    const id = String(casesData[i][COL_CASE_ID]).trim();
    if (id) caseMap[id] = casesData[i];
  }

  const now = new Date();
  const appsData = appsSheet.getDataRange().getValues();

  for (let i = 1; i < appsData.length; i++) {
    const row = appsData[i];

    const appId    = String(row[COL_APP_ID]           || '').trim();
    const userId   = String(row[COL_APP_USER_ID]      || '').trim();
    const status   = String(row[COL_APP_STATUS]       || '').trim();
    const sentAt   = row[COL_APP_CHECKIN_SENT];

    if (!appId || !userId)      continue;
    if (status !== 'approved')  continue;
    if (sentAt)                 continue; // 送付済み（二重送付防止）

    const startDt = parseStartDatetime(row[COL_APP_DATE], row[COL_APP_TIME]);
    if (!startDt) continue;

    // 開始10分前〜開始60分後の範囲内のみ送付
    const diffMin = (startDt - now) / 60000;
    if (diffMin > 10 || diffMin < -60) continue;

    const liffUrl = LIFF_BASE_URL + '?app_id=' + encodeURIComponent(appId);
    const sent = sendCheckinPush(userId, liffUrl);

    if (sent) {
      appsSheet.getRange(i + 1, COL_APP_CHECKIN_SENT + 1).setValue(now);
      console.log('[dispatch] sent to ' + userId + ', app=' + appId);
    }
  }
}

// ----  日時パース ----
function parseStartDatetime(reqDate, reqTime) {
  try {
    let dateStr;
    if (reqDate instanceof Date) {
      dateStr = Utilities.formatDate(reqDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      dateStr = String(reqDate).trim().slice(0, 10);
    }
    const timeStr = String(reqTime).trim().slice(0, 5);
    if (!dateStr || !timeStr) return null;
    return new Date(dateStr + 'T' + timeStr + ':00+09:00');
  } catch (e) {
    return null;
  }
}

// ---- LINE Push 送信 ----
function sendCheckinPush(userId, liffUrl) {
  try {
    const payload = {
      to: userId,
      messages: [
        {
          type: 'text',
          text: 'まもなく開始です。お店に着いたらチェックインしてください。',
        },
        {
          type: 'template',
          altText: 'チェックインはこちら',
          template: {
            type: 'buttons',
            text: '下のボタンからチェックインしてください',
            actions: [{ type: 'uri', label: 'チェックイン', uri: liffUrl }],
          },
        },
      ],
    };

    const res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    if (res.getResponseCode() !== 200) {
      console.error('[sendCheckinPush] HTTP ' + res.getResponseCode() + ': ' + res.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    console.error('[sendCheckinPush] error:', e);
    return false;
  }
}


// ============================================================
// ② doGet に追加: action=get_app
// ============================================================
// 既存の doGet の分岐に以下を追加:
//
//   case 'get_app':
//     return respond(getApp(params.app_id));
//
function getApp(appId) {
  try {
    if (!appId) return { success: false, error: 'app_id が指定されていません' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appsSheet  = ss.getSheetByName('applications');
    const casesSheet = ss.getSheetByName('cases');
    const appsData   = appsSheet.getDataRange().getValues();

    // 予約行を検索
    let appRow = null;
    for (let i = 1; i < appsData.length; i++) {
      if (String(appsData[i][COL_APP_ID]).trim() === String(appId).trim()) {
        appRow = appsData[i];
        break;
      }
    }
    if (!appRow) return { success: false, error: '予約が見つかりません' };

    // 案件情報を取得
    const caseId = String(appRow[COL_APP_CASE_ID]).trim();
    const casesData = casesSheet.getDataRange().getValues();
    let caseTitle = caseId;
    let caseArea  = '';
    for (let i = 1; i < casesData.length; i++) {
      if (String(casesData[i][COL_CASE_ID]).trim() === caseId) {
        caseTitle = String(casesData[i][COL_CASE_TITLE] || caseId);
        caseArea  = String(casesData[i][COL_CASE_AREA]  || '');
        break;
      }
    }

    // 日付を文字列化
    const reqDate = appRow[COL_APP_DATE];
    let dateStr = '';
    if (reqDate instanceof Date) {
      dateStr = Utilities.formatDate(reqDate, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      dateStr = String(reqDate).trim().slice(0, 10);
    }

    const checkinAt = appRow[COL_APP_CHECKIN_AT];

    return {
      success: true,
      app: {
        app_id:         appId,
        case_id:        caseId,
        case_title:     caseTitle,
        case_area:      caseArea,
        display_name:   String(appRow[COL_APP_DISPLAY_NAME]),
        requested_date: dateStr,
        requested_time: String(appRow[COL_APP_TIME]).trim().slice(0, 5),
        status:         String(appRow[COL_APP_STATUS]).trim(),
        checkin_at:     checkinAt ? String(checkinAt) : null,
      },
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}


// ============================================================
// ③ doPost に追加: action=checkin
// ============================================================
// 既存の doPost の分岐に以下を追加:
//
//   case 'checkin':
//     return respond(handleCheckin(payload));
//
function handleCheckin(payload) {
  try {
    const appId     = String(payload.app_id || '').trim();
    const inputCode = String(payload.code   || '').trim();

    if (!appId || !inputCode) {
      return { success: false, error: 'パラメータが不足しています' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const appsSheet  = ss.getSheetByName('applications');
    const casesSheet = ss.getSheetByName('cases');
    const appsData   = appsSheet.getDataRange().getValues();

    // 1. 予約を検索
    let appRow = null;
    let rowNumber = -1;
    for (let i = 1; i < appsData.length; i++) {
      if (String(appsData[i][COL_APP_ID]).trim() === appId) {
        appRow    = appsData[i];
        rowNumber = i + 1; // getRange は 1-indexed
        break;
      }
    }
    if (!appRow) return { success: false, error: '予約が見つかりません' };

    // 2. status=approved か
    if (String(appRow[COL_APP_STATUS]).trim() !== 'approved') {
      return { success: false, error: 'この予約はチェックイン対象外です' };
    }

    // 3. チェックイン済み確認
    if (appRow[COL_APP_CHECKIN_AT]) {
      return { success: false, error: 'すでにチェックイン済みです' };
    }

    // 4. 受付時間範囲確認（開始10分前〜開始60分後）
    const startDt = parseStartDatetime(appRow[COL_APP_DATE], appRow[COL_APP_TIME]);
    if (startDt) {
      const diffMin = (startDt - new Date()) / 60000;
      if (diffMin > 10) {
        return { success: false, error: 'チェックイン受付はまだ始まっていません（開始10分前から）' };
      }
      if (diffMin < -60) {
        return { success: false, error: 'チェックイン受付時間を過ぎています' };
      }
    }

    // 5. コード照合
    const caseId = String(appRow[COL_APP_CASE_ID]).trim();
    const casesData = casesSheet.getDataRange().getValues();
    let storedCode = '';
    for (let i = 1; i < casesData.length; i++) {
      if (String(casesData[i][COL_CASE_ID]).trim() === caseId) {
        storedCode = String(casesData[i][COL_CASE_CHECKIN_CODE] || '').trim();
        break;
      }
    }

    if (!storedCode) {
      return { success: false, error: 'チェックインコードが設定されていません（管理者へ連絡）' };
    }
    if (inputCode !== storedCode) {
      return { success: false, error: 'コードが違います' };
    }

    // 6. チェックイン記録
    const now = new Date();
    appsSheet.getRange(rowNumber, COL_APP_CHECKIN_AT + 1).setValue(now);

    // 7. 会場通知
    const displayName = String(appRow[COL_APP_DISPLAY_NAME]).trim();
    notifyVenueCheckin(caseId, displayName, appRow[COL_APP_DATE], appRow[COL_APP_TIME]);

    return { success: true, message: 'チェックイン完了！' };

  } catch (e) {
    console.error('[handleCheckin]', e);
    return { success: false, error: e.message };
  }
}

// ---- 会場通知（案件ごとの通知先設定が必要） ----
function notifyVenueCheckin(caseId, displayName, reqDate, reqTime) {
  // TODO: cases シートに venue_user_id 列を追加し、
  //       その店のLINEアカウントへ push する
  //
  // const venueUserId = getVenueUserId(caseId);
  // if (venueUserId) {
  //   sendLineTextPush(venueUserId,
  //     displayName + 'さんがチェックインしました（' + reqDate + ' ' + reqTime + '）');
  // }
  console.log('[checkin] ' + displayName + ' checked in: case=' + caseId
              + ' ' + reqDate + ' ' + reqTime);
}

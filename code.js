// ==========================================
// 富山市消防音楽隊 創立45周年記念演奏会
// 受付管理システム バックエンド処理 (Ver 1.2.4)
// (Auto-sync with clasp enabled)
// ==========================================

const SHEET_NAME = 'フォーム返答';
const CAPACITY_LIMIT = 2000;
const SYSTEM_VERSION = '1.2.4';

/**
 * Webアプリの初期表示（HTMLの提供）
 */
function doGet(e) {
  const action = e.parameter.action;

  // GitHub Pages 等の外部WebアプリからのAPIリクエスト処理
  if (action === 'search') {
    const code = e.parameter.code;
    const result = searchByCode(code);
    return createJsonResponse(result, e.parameter.callback);
  } else if (action === 'checkin') {
    const code = e.parameter.code;
    const actualCount = Number(e.parameter.count) || 1;
    const result = checkInUser(code, actualCount);
    return createJsonResponse(result, e.parameter.callback);
  }

  const page = e.parameter.p || 'form';
  if (page === 'admin') {
    const template = HtmlService.createTemplateFromFile('admin');
    template.version = SYSTEM_VERSION;
    return template.evaluate()
      .setTitle('受付用 照合システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  const template = HtmlService.createTemplateFromFile('form');
  template.version = SYSTEM_VERSION;
  return template.evaluate()
    .setTitle('富山市消防音楽隊 創立45周年記念演奏会 事前お申し込み')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * JSON または JSONP レスポンスを生成
 */
function createJsonResponse(data, callback) {
  const jsonStr = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonStr + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * シート取得処理
 */
function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.getSheets()[0];
  }
  return sheet;
}

/**
 * 新規お申し込み登録処理
 * スプレッドシート構造: 
 * A:受付番号 / B:来場者区分 / C:ニックネーム / D:人数 / E:ステータス / F:入場人数 / G:日時
 */
function registerNewUser(category, nickname, count) {
  const lock = LockService.getScriptLock();
  try {
    // 同時書き込みを防ぐため最大30秒待機
    lock.waitLock(30000);
  } catch (e) {
    return {
      success: false,
      message: 'アクセスが集中しています。しばらく経ってから再度お試しください。'
    };
  }

  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    
    const cleanNickname = normalizeString(nickname);
    if (!cleanNickname) {
      return {
        success: false,
        message: 'ニックネームを入力してください。'
      };
    }

    let currentTotalCount = 0;
    
    // 重複チェック & 現在の合計人数計算
    for (let i = 5; i < data.length; i++) {
      const rowNickname = String(data[i][2]);      // C列: ニックネーム
      const rowCount = Number(data[i][3]) || 0;    // D列: 人数
      const rowStatus = String(data[i][4]).trim(); // E列: ステータス
      
      if (rowStatus !== 'キャンセル') {
        // 同一ニックネームの重複チェック（全角半角・大文字小文字の表記揺れを吸収）
        if (normalizeString(rowNickname) === cleanNickname) {
          return {
            success: false,
            message: 'このニックネームはすでにお申し込みされています。別のニックネームをご使用ください。'
          };
        }
        currentTotalCount += rowCount;
      }
    }
    
    // 定員オーバーチェック
    const requestCount = Number(count);
    if (currentTotalCount + requestCount > CAPACITY_LIMIT) {
      return {
        success: false,
        message: '申し訳ありません。定員に達したため、お申し込みを受け付けられません。'
      };
    }

    // 5桁のユニークな受付番号を発行
    let code = generateUniqueCode(data);
    const timestamp = new Date();
    
    // スプレッドシートに書き込み (実シートの列順: A=受付番号, B=区分, C=ニックネーム, D=人数, E=ステータス, F=入場人数, G=日時)
    sheet.appendRow([
      code,
      category,
      nickname.trim(),
      requestCount,
      '未受付',
      '',
      timestamp
    ]);

    return {
      success: true,
      code: code
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  } finally {
    // ロックを確実に解放
    lock.releaseLock();
  }
}

/**
 * キャンセル処理（受付番号＋ニックネームの2重照合）
 */
function cancelUser(code, nickname) {
  const lock = LockService.getScriptLock();
  try {
    // 同時書き込みを防ぐため最大30秒待機
    lock.waitLock(30000);
  } catch (e) {
    return {
      success: false,
      message: 'アクセスが集中しています。しばらく経ってから再度お試しください。'
    };
  }

  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    const strCode = String(code).trim();
    const cleanNickname = normalizeString(nickname);

    if (!cleanNickname) {
      return {
        success: false,
        message: 'お申し込み時のニックネームを入力してください。'
      };
    }

    for (let i = 5; i < data.length; i++) {
      const rowCode = String(data[i][0]).trim();    // A列: 受付番号
      const rowNickname = String(data[i][2]);       // C列: ニックネーム
      const rowStatus = String(data[i][4]).trim();  // E列: ステータス

      // 受付番号とニックネームの両方を照合
      if (rowCode === strCode && normalizeString(rowNickname) === cleanNickname) {
        if (rowStatus === 'キャンセル') {
          return {
            success: false,
            message: 'このお申し込みはすでにキャンセルされています。'
          };
        }

        // E列（5列目）のステータスを「キャンセル」に更新
        sheet.getRange(i + 1, 5).setValue('キャンセル');

        return {
          success: true,
          message: 'お申し込みのキャンセルが完了いたしました。ご利用ありがとうございました。'
        };
      }
    }

    return {
      success: false,
      message: '受付番号またはニックネームが一致しません。入力内容をご確認ください。'
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  } finally {
    // ロックを確実に解放
    lock.releaseLock();
  }
}

/**
 * 受付用番号検索処理
 */
function searchByCode(code) {
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    const strCode = String(code).trim();

    for (let i = 5; i < data.length; i++) {
      const rowCode = String(data[i][0]).trim(); // A列: 受付番号

      if (rowCode === strCode) {
        return {
          success: true,
          code: data[i][0],        // A列: 受付番号
          category: data[i][1],    // B列: 来場者区分
          nickname: data[i][2],    // C列: ニックネーム
          count: data[i][3],       // D列: 人数
          status: data[i][4],      // E列: ステータス
          actualCount: data[i][5]  // F列: 入場人数
        };
      }
    }

    return {
      success: false,
      message: '該当する受付番号が存在しません。'
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  }
}

/**
 * チェックイン（入場確定）処理
 */
function checkInUser(code, actualCount) {
  const lock = LockService.getScriptLock();
  try {
    // 同時書き込みを防ぐため最大30秒待機
    lock.waitLock(30000);
  } catch (e) {
    return {
      success: false,
      message: 'アクセスが集中しています。しばらく経ってから再度お試しください。'
    };
  }

  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    const strCode = String(code).trim();

    for (let i = 5; i < data.length; i++) {
      const rowCode = String(data[i][0]).trim(); // A列: 受付番号

      if (rowCode === strCode) {
        sheet.getRange(i + 1, 6).setValue(actualCount); // F列（6列目）: 入場人数
        sheet.getRange(i + 1, 5).setValue('受付済');   // E列（5列目）: ステータス

        return {
          success: true
        };
      }
    }

    return {
      success: false,
      message: '更新対象のデータが見つかりませんでした。'
    };

  } catch (error) {
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  } finally {
    // ロックを確実に解放
    lock.releaseLock();
  }
}

/**
 * 5桁の被らないランダム数値を生成
 */
function generateUniqueCode(data) {
  const existingCodes = new Set();
  for (let i = 5; i < data.length; i++) {
    existingCodes.add(String(data[i][0]).trim()); // A列: 受付番号
  }

  let newCode = '';
  do {
    newCode = String(Math.floor(10000 + Math.random() * 90000));
  } while (existingCodes.has(newCode));

  return newCode;
}

/**
 * 文字列の表記揺れ吸収（全角半角・大文字小文字・前後空白の標準化）
 */
function normalizeString(str) {
  if (!str) return '';
  return String(str)
    .trim()
    .toLowerCase()
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });
}

/**
 * 【初回セットアップ用】スプレッドシート上部に集計ダッシュボードを自動構築する
 * ※ この関数はGASエディタから手動で1回だけ実行してください。
 *   既存のヘッダー行(1行目)の上に4行を挿入し、
 *   集計ラベル・数式・書式を自動設定します。
 *   実行後は「表示 > 固定 > 5行」を手動で設定してください。
 */
function setupDashboard() {
  const sheet = getTargetSheet();
  
  // ---- 上に4行挿入（既存ヘッダーが5行目、データが6行目～になる） ----
  sheet.insertRowsBefore(1, 4);
  
  // ==========================
  // 1行目: 全体集計ラベル
  // ==========================
  const row1Labels = ['定員', '総申込数', '残席数', '総入場数', '入場率', 'キャンセル数'];
  sheet.getRange(1, 1, 1, row1Labels.length).setValues([row1Labels]);
  sheet.getRange(1, 1, 1, row1Labels.length)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setBackground('#f0ebe0')
    .setFontColor('#5a4a3a');
  
  // ==========================
  // 2行目: 全体集計の数式・値
  // ==========================
  sheet.getRange('A2').setValue(CAPACITY_LIMIT);                            // 定員
  sheet.getRange('B2').setFormula('=SUMIFS(D6:D, E6:E, "<>キャンセル")');   // 総申込数
  sheet.getRange('C2').setFormula('=A2-B2');                                // 残席数
  sheet.getRange('D2').setFormula('=SUM(F6:F)');                            // 総入場数
  sheet.getRange('E2').setFormula('=IF(B2>0, D2/B2, 0)');                   // 入場率
  sheet.getRange('F2').setFormula('=COUNTIF(E6:E, "キャンセル")');          // キャンセル数
  
  // 2行目の書式設定
  sheet.getRange(2, 1, 1, 6)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.getRange('E2').setNumberFormat('0.0%');  // 入場率を%表示
  
  // ==========================
  // 3行目: 区分別集計ラベル
  // ==========================
  const row3Labels = ['一般(申込)', '一般(入場)', '職員(申込)', '職員(入場)', '団員(申込)', '団員(入場)', '関係者(申込)', '関係者(入場)'];
  sheet.getRange(3, 1, 1, row3Labels.length).setValues([row3Labels]);
  sheet.getRange(3, 1, 1, row3Labels.length)
    .setFontWeight('bold')
    .setFontSize(9)
    .setHorizontalAlignment('center')
    .setBackground('#f0ebe0')
    .setFontColor('#5a4a3a');
  
  // ==========================
  // 4行目: 区分別集計の数式
  // ==========================
  sheet.getRange('A4').setFormula('=SUMIFS(D6:D, B6:B, "一般", E6:E, "<>キャンセル")');
  sheet.getRange('B4').setFormula('=SUMIFS(F6:F, B6:B, "一般")');
  sheet.getRange('C4').setFormula('=SUMIFS(D6:D, B6:B, "消防職員", E6:E, "<>キャンセル")');
  sheet.getRange('D4').setFormula('=SUMIFS(F6:F, B6:B, "消防職員")');
  sheet.getRange('E4').setFormula('=SUMIFS(D6:D, B6:B, "消防団員", E6:E, "<>キャンセル")');
  sheet.getRange('F4').setFormula('=SUMIFS(F6:F, B6:B, "消防団員")');
  sheet.getRange('G4').setFormula('=SUMIFS(D6:D, B6:B, "関係者", E6:E, "<>キャンセル")');
  sheet.getRange('H4').setFormula('=SUMIFS(F6:F, B6:B, "関係者")');
  
  // 4行目の書式設定
  sheet.getRange(4, 1, 1, 8)
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // ==========================
  // 5行目: データヘッダー行の書式整え
  // ==========================
  sheet.getRange(5, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground('#e8e0d0')
    .setHorizontalAlignment('center');
  
  // ==========================
  // 罫線の設定
  // ==========================
  // 1-2行目（全体集計エリア）に外枠
  sheet.getRange(1, 1, 2, 6).setBorder(true, true, true, true, null, null, '#8b7355', SpreadsheetApp.BorderStyle.SOLID);
  // 3-4行目（区分別集計エリア）に外枠
  sheet.getRange(3, 1, 2, 8).setBorder(true, true, true, true, null, null, '#8b7355', SpreadsheetApp.BorderStyle.SOLID);
  // 5行目（ヘッダー行）に下線
  sheet.getRange(5, 1, 1, 7).setBorder(null, null, true, null, null, null, '#8b7355', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  SpreadsheetApp.flush();
  Logger.log('✅ ダッシュボードのセットアップが完了しました。「表示 > 固定 > 5行」を手動で設定してください。');
}
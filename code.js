// ==========================================
// 富山市消防音楽隊 創立45周年記念演奏会
// 受付管理システム バックエンド処理 (Ver 1.1)
// ==========================================

const SHEET_NAME = 'フォーム返答';
const CAPACITY_LIMIT = 500;

/**
 * Webアプリの初期表示（HTMLの提供）
 */
function doGet(e) {
  const page = e.parameter.p || 'form';
  if (page === 'admin') {
    return HtmlService.createTemplateFromFile('admin')
      .evaluate()
      .setTitle('受付用 照合システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }
  return HtmlService.createTemplateFromFile('form')
    .evaluate()
    .setTitle('富山市消防音楽隊 創立45周年記念演奏会 事前お申し込み')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
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
    for (let i = 1; i < data.length; i++) {
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
  }
}

/**
 * キャンセル処理（受付番号＋ニックネームの2重照合）
 */
function cancelUser(code, nickname) {
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

    for (let i = 1; i < data.length; i++) {
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

    for (let i = 1; i < data.length; i++) {
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
  try {
    const sheet = getTargetSheet();
    const data = sheet.getDataRange().getValues();
    const strCode = String(code).trim();

    for (let i = 1; i < data.length; i++) {
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
  }
}

/**
 * 5桁の被らないランダム数値を生成
 */
function generateUniqueCode(data) {
  const existingCodes = new Set();
  for (let i = 1; i < data.length; i++) {
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
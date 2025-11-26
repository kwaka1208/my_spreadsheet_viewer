// ▼▼▼ 設定エリア ▼▼▼
// ユーザーのマイドライブにこの名前で設定ファイルが作られます
const CONFIG_FILE_NAME = "SpreadsheetViewer_Settings";
const CONFIG_SHEET_NAME = "Settings";
// ▲▲▲ 設定ここまで ▲▲▲

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('My スプレッドシート Viewer') // ← ここを変更
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 設定用スプレッドシートを取得（なければ作成）するヘルパー関数
 */
function getOrCreateConfigSheet_() {
  try {
    // 1. マイドライブから同名のファイルを検索 (マイファイルのみ、ゴミ箱除く)
    const files = DriveApp.getFilesByName(CONFIG_FILE_NAME);
    
    let ss;
    if (files.hasNext()) {
      // 見つかったらそれを使う
      const file = files.next();
      ss = SpreadsheetApp.open(file);
    } else {
      // 見つからなければ新規作成
      ss = SpreadsheetApp.create(CONFIG_FILE_NAME);
    }

    // 2. シートの取得または作成
    let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
    if (!sheet) {
      // なければ作成
      sheet = ss.insertSheet(CONFIG_SHEET_NAME);
      // 古いシート(シート1)が残っている場合があるので整理しても良いが、今回はそのまま
    }

    // 3. ヘッダー行がない（新規作成直後）なら追加
    if (sheet.getLastRow() === 0) {
      // [表示名, URL, シート名, 見出し行数, 表示列]
      sheet.appendRow(["表示名", "URL", "シート名", "見出し行数", "表示列"]);
    }

    return sheet;

  } catch (e) {
    // DriveAppの権限がない場合などが考えられる
    throw new Error("設定ファイルの取得に失敗しました: " + e.message);
  }
}

/**
 * 設定リストを取得する関数
 */
function getSheetList() {
  const settings = getSettingsFromSheet_();
  
  return settings.map((item, index) => {
    return { 
      index: index, 
      name: item.name,
      headerRows: item.headerRows
    };
  });
}

/**
 * 内部処理: シートから設定データを読み込む
 */
function getSettingsFromSheet_() {
  try {
    const sheet = getOrCreateConfigSheet_(); // ここで取得・作成
    const rows = sheet.getDataRange().getDisplayValues();
    
    // 1行目はヘッダーなので削除
    if (rows.length > 0) rows.shift();

    return rows.map(row => {
      // 列指定のパース
      let cols = [];
      if (row[4]) {
        cols = row[4].toString().replace(/，/g, ',').split(',').map(s => s.trim());
      }

      return {
        name: row[0],
        url: row[1],
        sheetName: row[2],
        headerRows: parseInt(row[3], 10) || 1,
        visibleColumns: cols
      };
    }).filter(item => item.name && item.url);

  } catch (e) {
    console.error(e);
    return [];
  }
}

/**
 * 設定を追加する関数
 */
function addSetting(form) {
  try {
    const sheet = getOrCreateConfigSheet_(); // ここで取得・作成
    
    if (!form.name || !form.url || !form.sheetName) {
      throw new Error("必須項目が不足しています");
    }

    const newRow = [
      form.name,
      form.url,
      form.sheetName,
      form.headerRows || 1,
      form.visibleColumns || ""
    ];

    sheet.appendRow(newRow);
    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * データ取得メイン関数
 */
function getSheetData(request) {
  try {
    let url = "";
    let sheetName = "";
    let rawColumns = [];

    if (request.type === 'preset') {
      const settings = getSettingsFromSheet_();
      const setting = settings[request.index];
      if (!setting) throw new Error("設定が見つかりません。");
      
      url = setting.url;
      sheetName = setting.sheetName;
      rawColumns = setting.visibleColumns || [];
    } 
    else if (request.type === 'manual') {
      url = request.url;
      sheetName = request.sheetName;
      rawColumns = request.visibleColumns || [];
    }

    // 列変換
    let targetColumns = [];
    if (rawColumns && rawColumns.length > 0) {
      targetColumns = rawColumns.map(convertColumnToNumber_).filter(n => n !== null && n > 0);
    }

    const ss = SpreadsheetApp.openByUrl(url);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('指定されたシートが見つかりませんでした。');
    
    const rawData = sheet.getDataRange().getDisplayValues();
    if (rawData.length === 0) throw new Error('シートにデータがありません。');
    
    // フィルタリング
    if (targetColumns.length === 0) {
      return { success: true, data: rawData };
    }

    const filteredData = rawData.map(row => {
      return targetColumns.map(colNum => {
        return row[colNum - 1] || "";
      });
    });

    return { success: true, data: filteredData };
    
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ヘルパー: 列文字を数値に変換
function convertColumnToNumber_(val) {
  if (!val) return null;
  if (typeof val === 'number') return val;
  const str = val.toString().trim().toUpperCase();
  if (/^\d+$/.test(str)) return parseInt(str, 10);
  let num = 0;
  for (let i = 0; i < str.length; i++) {
    num = num * 26 + (str.charCodeAt(i) - 64);
  }
  return num;
}
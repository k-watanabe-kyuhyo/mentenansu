/**
 * 設定: シート名
 */
const SHEET_NAME = 'maintenance_records';

/**
 * Webアプリの表示 (doGet)
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('機器メンテ管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * スプレッドシートの取得
 */
function getSpreadsheet() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  
  if (!spreadsheetId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    scriptProperties.setProperty('SPREADSHEET_ID', ss.getId());
    return ss;
  }
  
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * 全レコードを取得する
 */
function getMaintenanceRecords() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return [];
   
    const range = sheet.getDataRange();
    const data = range.getValues();
    if (data.length <= 1) return [];
    data.shift();
   
    return data.filter(row => row[0]).map((row) => {
      // ドライブURLを表示可能な直接リンクに変換する
      const formatDriveUrl = (url) => {
        if (!url || typeof url !== 'string') return '';
        const idMatch = url.match(/id=([a-zA-Z0-9_-]+)/) || url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (idMatch && idMatch[1]) {
          // 最も表示互換性の高いGoogleユーザーコンテンツURL形式に変換
          return "https://lh3.googleusercontent.com/u/0/d/" + idMatch[1];
        }
        return url;
      };

      const buildPhotosArray = (...urls) => {
        return urls
          .filter(url => url && typeof url === 'string' && url.trim() !== '')
          .map(url => ({ data: formatDriveUrl(url) }));
      };

      return {
        id: String(row[0] || ''),
        createdAt: row[1] instanceof Date ? Utilities.formatDate(row[1], "JST", "yyyy-MM-dd HH:mm:ss") : String(row[1] || ''),
        date: row[2] instanceof Date ? Utilities.formatDate(row[2], "JST", "yyyy-MM-dd") : String(row[2] || ''),
        locationId: String(row[3] || ''),
        machineName: String(row[4] || ''),
        repairCost: isNaN(row[5]) || row[5] === '' ? 0 : Number(row[5]),
        repairTime: isNaN(row[6]) || row[6] === '' ? 0 : Number(row[6]),
        operator: String(row[7] || ''),
        repairDetails: String(row[8] || ''),
        machinePhotosJson: buildPhotosArray(row[10], row[11]),
        faultsJson: buildPhotosArray(row[12], row[13], row[14], row[15]),
        quotesJson: buildPhotosArray(row[16], row[17])
      };
    });
  } catch (e) {
    console.error('取得エラー: ' + e.toString());
    return [];
  }
}

/**
 * レコードを保存（新規・更新）する
 */
function saveRecord(record) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = ['ID', '作成日時', '日付', '場所ID', '機器名', '修理費用', '修理時間(h)', '入力者', '修理詳細', '故障内容テキスト', '外観1', '外観2', '故障1', '故障2', '故障3', '故障4', '見積1', '見積2'];
      sheet.appendRow(headers);
    }

    const data = sheet.getDataRange().getValues();
    let targetRow = -1;
    const currentId = record.id || 'ID_' + Utilities.getUuid();
    const createdAt = record.createdAt || Utilities.formatDate(new Date(), "JST", "yyyy-MM-dd HH:mm:ss");

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(currentId)) {
        targetRow = i + 1;
        break;
      }
    }

    function processPhotos(photosArray, maxCount) {
      const urls = [];
      let items = photosArray;
      if (typeof items === 'string') {
        try { items = JSON.parse(items); } catch(e) { items = []; }
      }
      if (!Array.isArray(items)) items = [];

      for (let i = 0; i < maxCount; i++) {
        if (i < items.length && items[i]) {
          const item = items[i];
          const dataUrl = item.data || item;
          if (typeof dataUrl === 'string') {
            if (dataUrl.indexOf('data:') === 0 && dataUrl.indexOf(';base64,') > -1) {
              urls.push(saveFileToDrive(dataUrl));
            } else {
              urls.push(dataUrl);
            }
          } else {
            urls.push('');
          }
        } else {
          urls.push('');
        }
      }
      return urls;
    }

    const mainPhotos = processPhotos(record.machinePhotosJson, 2);
    const faultPhotos = processPhotos(record.faultsJson, 4);
    const quotePhotos = processPhotos(record.quotesJson, 2);

    const rowData = [
      currentId,
      createdAt,
      record.date || '',
      record.locationId || '',
      record.machineName || '',
      record.repairCost || 0,
      record.repairTime || 0,
      record.operator || '',
      record.repairDetails || '',
      '',
      mainPhotos[0],
      mainPhotos[1],
      faultPhotos[0],
      faultPhotos[1],
      faultPhotos[2],
      faultPhotos[3],
      quotePhotos[0],
      quotePhotos[1]
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    SpreadsheetApp.flush();
    return getMaintenanceRecords();
  } catch (e) {
    console.error('保存エラー: ' + e.toString());
    throw new Error('保存に失敗しました: ' + e.message);
  }
}

/**
 * DataURL (Base64) をGoogleドライブに保存しURLを返す
 */
function saveFileToDrive(dataUrl) {
  try {
    const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
    let folder;
    if (folderId) {
      try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = DriveApp.getRootFolder(); }
    } else {
      folder = DriveApp.getRootFolder();
    }

    const contentType = dataUrl.substring(5, dataUrl.indexOf(';'));
    const bytes = Utilities.base64Decode(dataUrl.substring(dataUrl.indexOf('base64,') + 7));
    const fileName = "repair_" + Utilities.formatDate(new Date(), "JST", "yyyyMMdd_HHmmss") + "_" + Math.floor(Math.random() * 1000);
    
    let extension = ".jpg";
    if (contentType === "image/png") extension = ".png";
    else if (contentType === "application/pdf") extension = ".pdf";

    const file = folder.createFile(Utilities.newBlob(bytes, contentType, fileName + extension));
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 保存時は従来のID付きURL
    return "https://drive.google.com/uc?id=" + file.getId();
  } catch (e) {
    console.error('ドライブ保存エラー: ' + e.toString());
    return "";
  }
}

/**
 * レコードを削除する
 */
function deleteRecord(id) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
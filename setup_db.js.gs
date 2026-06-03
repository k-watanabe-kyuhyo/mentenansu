/**
 * データベース（スプレッドシート）初期セットアップ用スクリプト
 * * 変更点: 写真を1枚ずつ個別のセルに保存する構成に変更
 */

function setupSpreadsheetDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "maintenance_records";
 
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      `シート "${sheetName}" は既に存在します。\n再セットアップを行うと列の構成が変更されます（既存のJSON形式の写真は表示されなくなる可能性があります）。\n続行しますか？`,
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;
    sheet.activate();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // 2. カラム定義（合計17カラム）
  const columns = [
    { key: "id", name: "ID", width: 250 },
    { key: "createdAt", name: "作成日時", width: 150 },
    { key: "date", name: "日付", width: 100 },
    { key: "locationId", name: "場所ID", width: 100 },
    { key: "machineName", name: "機器名", width: 200 },
    { key: "repairCost", name: "修理費用", width: 100 },
    { key: "repairTime", name: "修理時間(h)", width: 80 },
    { key: "operator", name: "入力者", width: 120 },
    { key: "repairDetails", name: "修理詳細", width: 300 },
    { key: "faultsJson", name: "故障内容テキスト", width: 150 }, // 互換性のために残す（基本は空）
    // 個別写真セル
    { key: "photo_main_1", name: "外観1", width: 100 },
    { key: "photo_main_2", name: "外観2", width: 100 },
    { key: "photo_fault_1", name: "故障1", width: 100 },
    { key: "photo_fault_2", name: "故障2", width: 100 },
    { key: "photo_fault_3", name: "故障3", width: 100 },
    { key: "photo_fault_4", name: "故障4", width: 100 },
    { key: "photo_quote_1", name: "見積1", width: 100 },
    { key: "photo_quote_2", name: "見積2", width: 100 }
  ];

  const headerRowIndex = 1;
  const lastColIndex = columns.length;

  const headerNames = columns.map(c => c.name);
  const headerKeys = columns.map(c => c.key);
 
  const range = sheet.getRange(headerRowIndex, 1, 1, lastColIndex);
  range.setValues([headerNames]);
  range.setNotes([headerKeys]);

  // 装飾
  range.setBackground("#4285F4").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
  columns.forEach((col, index) => sheet.setColumnWidth(index + 1, col.width));
  sheet.setFrozenRows(1);

  // フォーマット
  sheet.getRange(2, 6, sheet.getMaxRows() - 1, 1).setNumberFormat("¥#,##0");
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd");

  SpreadsheetApp.getUi().alert("セットアップが完了しました。新しい列構成が適用されました。");
}
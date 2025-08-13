function buildAttackList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const attackSheet = ss.getSheetByName("アタックリスト");

  // 管理シートを開く（固定ID）
  const manageSS = SpreadsheetApp.openById("1oeNGXFXgxqKCrP86Tps8mFTg2K21hglpN8_DLOJqGfU");
  const manageSheet = manageSS.getSheetByName("List");

  // アタックリスト初期化
  attackSheet.clearContents();
  attackSheet.appendRow([
    "会社名",
    "支店・部署",
    "マンション名",
    "フロント担当者名",
    "提案可否",
    "次回理事会の関与形式",
    "次の理事会日/日付不明は1日で仮設定"
  ]);

  const cutoff = new Date("2025/07/16");
  const manageData = manageSheet.getRange(2,1,manageSheet.getLastRow()-1,4).getValues(); 
  // A:会社名, B:支店, C:sheet_id, D:タブ名

  manageData.forEach(row => {
    const [company, branch, sheetId, tabName] = row;
    if (!sheetId) return;

    const extSS = SpreadsheetApp.openById(sheetId);
    const oneOnOneSheet = extSS.getSheetByName(tabName);
    const data = oneOnOneSheet.getDataRange().getValues();
    const headers = data[0];

    // 列番号をヘッダー名から特定
    const mansionCol    = headers.indexOf("マンション名");
    const frontCol      = headers.indexOf("フロント担当者名");
    const teianCol      = headers.indexOf("提案可否");
    const kanyoCol      = headers.indexOf("次回理事会の関与形式");
    const dateCol       = headers.indexOf("次の理事会日/日付不明は1日で仮設定");
    const mailBeforeCol = headers.indexOf("理事会前メール送信日");
    const mailAfterCol  = headers.indexOf("理事会後メール送信日");
    const sfaFlagCol    = headers.indexOf("SFA商談化フラグ");

    // データ行ループ
    data.slice(1).forEach(r => {
      const mansion   = r[mansionCol];
      const teian     = r[teianCol];
      const sfaFlag   = r[sfaFlagCol];
      const mailBefore= r[mailBeforeCol];
      const mailAfter = r[mailAfterCol];
      const dateVal   = r[dateCol];

      // 条件チェック
      if (!mansion) return;
      if (teian !== "提案可") return;
      if (sfaFlag !== "リード化") return;
      if (mailBefore || mailAfter) return;
      if (!dateVal) return;

      const date = new Date(dateVal);
      if (date <= cutoff) {
        attackSheet.appendRow([
          company,
          branch,
          mansion,
          r[frontCol],
          teian,
          r[kanyoCol],
          date
        ]);
      }
    });
  });
}

// ============================================================
// 進階練習：一鍵生成客戶報價單
// 對應：Session 7（格式設定、文字/數字/日期格式）
// ============================================================

/**
 * 一鍵生成專業報價單
 */
function 生成報價單() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var 客戶表 = ss.getSheetByName("報價資料");
    if (!客戶表) { SpreadsheetApp.getUi().alert("❌ 請先初始化"); return; }

    var ui = SpreadsheetApp.getUi();
    
    // 1. 提示輸入客戶名稱
    var 客戶回應 = ui.prompt("📝 報價單", "請輸入客戶名稱：", ui.ButtonSet.OK_CANCEL);
    if (客戶回應.getSelectedButton() !== ui.Button.OK) return;
    var 客戶 = 客戶回應.getResponseText().trim() || "範例客戶";

    // 2. 提示輸入客戶 Email
    var Email回應 = ui.prompt("📧 客戶電子信箱", "請輸入收件人 Email 位址：", ui.ButtonSet.OK_CANCEL);
    if (Email回應.getSelectedButton() !== ui.Button.OK) return;
    var 客戶Email = Email回應.getResponseText().trim();
    if (!客戶Email || 客戶Email.indexOf("@") === -1) {
      ui.alert("❌ 請輸入有效的 Email 位址！已取消生成報價單。");
      return;
    }

    var 編號 = "QT-" + Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd") + "-" +
               String(Math.floor(Math.random() * 100)).padStart(3, "0");

    var sheet = ss.insertSheet(編號, 0);

    // ===== 報價單格式 =====

    // 公司 Logo 區
    sheet.getRange("A1:F1").merge();
    sheet.getRange("A1").setValue("ABC 科技股份有限公司")
      .setFontSize(20).setFontWeight("bold").setFontColor("#1a237e");
    sheet.setRowHeight(1, 50);

    sheet.getRange("A2:F2").merge();
    sheet.getRange("A2").setValue("台北市信義區信義路五段 7 號 ｜ Tel: 02-2345-6789 ｜ www.abc-tech.com")
      .setFontSize(9).setFontColor("#666");

    // 報價單標題
    sheet.getRange("A4:F4").merge();
    sheet.getRange("A4").setValue("📋 報 價 單")
      .setFontSize(24).setFontWeight("bold").setHorizontalAlignment("center")
      .setBackground("#1a237e").setFontColor("#fff");
    sheet.setRowHeight(4, 45);

    // 客戶資訊
    var 今天 = new Date();
    var 有效日 = new Date(今天); 有效日.setDate(今天.getDate() + 30);

    sheet.getRange(6, 1, 4, 6).setValues([
      ["報價編號", 編號, "", "客戶名稱", 客戶, ""],
      ["報價日期", Utilities.formatDate(今天, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡人", "", ""],
      ["有效期限", Utilities.formatDate(有效日, "Asia/Taipei", "yyyy/MM/dd"), "", "聯絡電話", "", ""],
      ["業務人員", "林冠廷", "", "傳真", "", ""]
    ]);

    for (var r = 6; r <= 9; r++) {
      sheet.getRange(r, 1).setFontWeight("bold").setBackground("#e8eaf6");
      sheet.getRange(r, 4).setFontWeight("bold").setBackground("#e8eaf6");
    }

    // 品項表格
    sheet.getRange("A11:F11").setValues([["項次", "品名/規格", "單位", "數量", "單價", "金額"]]);
    sheet.getRange("A11:F11").setBackground("#283593").setFontColor("#fff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setRowHeight(11, 35);

    // 讀取報價品項
    var 品項資料 = 客戶表.getDataRange().getValues();
    var 小計 = 0;

    for (var i = 1; i < 品項資料.length; i++) {
      var 列 = 11 + i;
      var 金額 = 品項資料[i][3] * 品項資料[i][4]; // 數量 × 單價
      小計 += 金額;

      sheet.getRange(列, 1, 1, 6).setValues([[
        i, 品項資料[i][0], 品項資料[i][1], 品項資料[i][3], 品項資料[i][4], 金額
      ]]);
      sheet.getRange(列, 1).setHorizontalAlignment("center");
      sheet.getRange(列, 4, 1, 3).setNumberFormat("#,##0");
      sheet.getRange(列, 4, 1, 3).setHorizontalAlignment("right");

      // 斑馬紋
      if (i % 2 === 0) sheet.getRange(列, 1, 1, 6).setBackground("#f5f5f5");
    }

    // 合計區
    var 稅金 = Math.round(小計 * 0.05);
    var 總計 = 小計 + 稅金;
    var 合計起始 = 12 + 品項資料.length - 1;

    var 合計資料 = [
      ["", "", "", "", "小　計", 小計],
      ["", "", "", "", "稅金 (5%)", 稅金],
      ["", "", "", "", "總　計", 總計]
    ];

    sheet.getRange(合計起始, 1, 3, 6).setValues(合計資料);
    sheet.getRange(合計起始, 5, 3, 1).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRange(合計起始, 6, 3, 1).setNumberFormat("NT$#,##0").setFontWeight("bold");
    sheet.getRange(合計起始 + 2, 5, 1, 2).setBackground("#e8eaf6")
      .setFontSize(14).setBorder(true, true, true, true, false, false,
        "#1a237e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 備註
    var 備註列 = 合計起始 + 5;
    sheet.getRange(備註列, 1, 1, 6).merge();
    sheet.getRange(備註列, 1).setValue("📌 備註：").setFontWeight("bold");
    sheet.getRange(備註列 + 1, 1, 1, 6).merge();
    sheet.getRange(備註列 + 1, 1).setValue(
      "1. 報價有效期限 30 天\n2. 付款條件：月結 30 天\n3. 交貨期：訂購後 7~14 個工作天\n4. 以上報價含安裝及教育訓練"
    ).setWrap(true).setFontColor("#555");
    sheet.setRowHeight(備註列 + 1, 80); // 調整備註高度以容納多行文字

    // 欄寬自適應調整
    // 暫時將 A 欄中會影響寬度計算的合併儲存格長文字清空，避免欄 1 (項次) 被異常拉寬
    var tempA1 = sheet.getRange("A1").getValue();
    var tempA2 = sheet.getRange("A2").getValue();
    var tempA4 = sheet.getRange("A4").getValue();
    var tempRemark1 = sheet.getRange(備註列, 1).getValue();
    var tempRemark2 = sheet.getRange(備註列 + 1, 1).getValue();

    sheet.getRange("A1").setValue("");
    sheet.getRange("A2").setValue("");
    sheet.getRange("A4").setValue("");
    sheet.getRange(備註列, 1).setValue("");
    sheet.getRange(備註列 + 1, 1).setValue("");

   // 進行第 1 至 6 欄的自適應寬度調整與防護限制
    for (var c = 1; c <= 6; c++) {
      sheet.autoResizeColumn(c);
      var width = sheet.getColumnWidth(c);
      
      // 依欄位屬性給予合理的最低寬度限制
      if (c === 1 && width < 100) sheet.setColumnWidth(c, 100);       // ✅ 修改這裡：將 A 欄下限改為 100，確保「報價編號」能完整顯示
      else if (c === 2 && width < 200) sheet.setColumnWidth(c, 200);  // 品名/規格
      else if (c === 3 && width < 60) sheet.setColumnWidth(c, 60);    // 單位
      else if (c === 4 && width < 80) sheet.setColumnWidth(c, 80);    // 數量
      else if (c === 5 && width < 100) sheet.setColumnWidth(c, 100);  // 單價
      else if (c === 6 && width < 120) sheet.setColumnWidth(c, 120);  // 金額
    }

    // 還原 A 欄長文字
    sheet.getRange("A1").setValue(tempA1);
    sheet.getRange("A2").setValue(tempA2);
    sheet.getRange("A4").setValue(tempA4);
    sheet.getRange(備註列, 1).setValue(tempRemark1);
    sheet.getRange(備註列 + 1, 1).setValue(tempRemark2);

    // 整體框線
    sheet.getRange(11, 1, 品項資料.length + 3, 6)
      .setBorder(true, true, true, true, true, true, "#bdbdbd", SpreadsheetApp.BorderStyle.SOLID);

    sheet.setFrozenRows(0);

    // ===== 匯出 PDF 並寄送 Email =====
    
    // 強制將所有未儲存的變更寫入工作表，以利匯出正確資料的 PDF
    SpreadsheetApp.flush();

    var ssId = ss.getId();
    var sheetId = sheet.getSheetId();

    // 建構 PDF 匯出 URL 與參數
    var url = ss.getUrl().replace(/edit$/, '') + 'export?';
    var exportOptions = {
      exportFormat: 'pdf',
      format: 'pdf',
      size: 'A4',             // A4 紙張
      portrait: 'true',       // 直向
      fitw: 'true',           // 符合寬度 (一頁寬)
      gridlines: 'false',     // 隱藏格線
      printtitle: 'false',    // 不重複列印標題
      sheetnames: 'false',    // 不重複列印工作表名稱
      fzr: 'false',           // 不重複列印凍結列
      gid: sheetId            // 指定只匯出該報價單
    };

    var urlParts = [];
    for (var key in exportOptions) {
      urlParts.push(key + '=' + encodeURIComponent(exportOptions[key]));
    }
    var exportUrl = url + urlParts.join('&');

    var response = UrlFetchApp.fetch(exportUrl, {
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("PDF 轉換失敗，HTTP 狀態碼：" + response.getResponseCode());
    }

    var blob = response.getBlob().setName(編號 + "_" + 客戶 + ".pdf");

    // 發送電子郵件並附加 PDF
    var 主旨 = "【ABC 科技】您的報價單已生成 - 編號：" + 編號;
    var 信件內文 = "親愛的 " + 客戶 + " 您好：\n\n" +
               "感謝您的諮詢！附件為為您量身打造的報價單（編號：" + 編號 + "），請查收。\n" +
               "如有任何問題，歡迎隨時回信與我們聯繫。\n\n" +
               "祝 順心\n" +
               "ABC 科技團隊 敬上";

    MailApp.sendEmail({
      to: 客戶Email,
      subject: 主旨,
      body: 信件內文,
      attachments: [blob]
    });

    SpreadsheetApp.getUi().alert("✅ 報價單 " + 編號 + " 已生成，且已轉成 PDF 並寄信至 " + 客戶Email + "！\n總金額：NT$ " + 總計.toLocaleString());

  } catch (錯誤) { Logger.log("❌ " + 錯誤.message); SpreadsheetApp.getUi().alert("❌ " + 錯誤.message); }
}

function 初始化報價資料() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("報價資料");
  if (!sheet) sheet = ss.insertSheet("報價資料"); else sheet.clear();

  var 標題 = [["品名/規格", "單位", "說明", "數量", "單價"]];
  var 資料 = [
    ["AI 智慧會議系統（基本版）", "套", "含語音辨識、自動會議紀要", 1, 180000],
    ["智慧文件管理模組", "授權", "OCR + AI 分類，10 人授權", 10, 12000],
    ["自動化報表工具", "授權", "每日/週/月報表自動產生", 5, 8500],
    ["教育訓練", "小時", "現場教育訓練（含教材）", 16, 3000],
    ["系統維護（年約）", "年", "含系統更新與技術支援", 1, 50000],
    ["客製化開發", "人天", "依需求客製功能開發", 10, 8000]
  ];

  sheet.getRange(1, 1, 1, 5).setValues(標題);
  sheet.getRange(2, 1, 資料.length, 5).setValues(資料);
  sheet.getRange("A1:E1").setBackground("#283593").setFontColor("#fff").setFontWeight("bold");
  sheet.getRange("E2:E7").setNumberFormat("#,##0");
  sheet.setFrozenRows(1);
  for (var c = 1; c <= 5; c++) sheet.autoResizeColumn(c);

  SpreadsheetApp.getUi().alert("✅ 報價資料已建立！請執行「生成報價單」。");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🤖 報價單系統")
    .addItem("📦 初始化報價資料", "初始化報價資料")
    .addItem("📋 生成報價單", "生成報價單")
    .addToUi();
}

/**
 * 部署步驟：
 * 1. 在這份 Google Sheet 內，點選「擴充功能」→「Apps Script」
 * 2. 刪除預設內容，貼上這段程式碼
 * 3. 點選「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我（你的帳號）
 *    - 誰可以存取：任何人
 * 4. 部署後複製「網頁應用程式網址」，貼到 index.html 裡的 GAS_URL
 *
 * 需要在同一份試算表中，另外建立一個名為「廠商名單」的工作表，
 * 欄位依序為：手機 / 廠商 / 姓名 / 對接同仁（第一列為標題列）。
 * 特約廠商登入時會用手機號碼查詢這張表，自動帶入廠商 / 姓名 / 對接同仁。
 *
 * --- Synology Chat 通知（選用）---
 * 1. Synology Chat 後台 →「整合」→「Incoming Webhook」→ 新增，複製 Webhook 網址
 *    （網址主機部分要用 QuickConnect 或 DDNS 對外網址，不能用內網 IP）
 * 2. 在這個 Apps Script 專案：左側齒輪圖示「專案設定」→「指令碼屬性」→ 新增屬性
 *    屬性名稱：SYNO_WEBHOOK_URL　值：貼上剛剛複製的 Webhook 網址
 *    （這樣 Token 只存在 Apps Script 裡，不會出現在 GitHub 原始碼中）
 * 3. 沒有設定這個屬性時，通知會自動略過，不影響登記本身。
 */

var VENDOR_SHEET_NAME = "廠商名單";

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  sheet.appendRow([
    data.date || "",
    data.vendor || "",
    data.name || "",
    data.arrival || "",
    data.contact || "",
  ]);

  notifySynologyChat(data);

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function notifySynologyChat(data) {
  var webhookUrl = PropertiesService.getScriptProperties().getProperty("SYNO_WEBHOOK_URL");
  if (!webhookUrl) {
    console.log("SYNO_WEBHOOK_URL 尚未設定，略過通知");
    return;
  }

  var message =
    "📋 新訪客登記\n" +
    "日期：" + (data.date || "") + "\n" +
    "廠商：" + (data.vendor || "") + "\n" +
    "姓名：" + (data.name || "") + "\n" +
    "預計抵達：" + (data.arrival || "") + "\n" +
    "對接同仁：" + (data.contact || "");

  try {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: "post",
      payload: { payload: JSON.stringify({ text: message }) },
      muteHttpExceptions: true,
    });
    console.log(
      "Synology webhook 回應碼：" + response.getResponseCode() +
      "，內容：" + response.getContentText()
    );
  } catch (err) {
    console.log("Synology webhook 呼叫失敗：" + err.message);
  }
}

function doGet(e) {
  var action = e.parameter.action;
  var result = { error: "invalid action" };

  if (action === "lookup") {
    result = lookupVendor(e.parameter.phone || "");
  }

  var callback = e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function lookupVendor(phone) {
  phone = String(phone).trim();
  if (!phone) return { found: false };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VENDOR_SHEET_NAME);
  if (!sheet) return { found: false };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === phone) {
      return {
        found: true,
        vendor: rows[i][1] || "",
        name: rows[i][2] || "",
        contact: rows[i][3] || "",
      };
    }
  }
  return { found: false };
}

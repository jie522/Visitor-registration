/**
 * 部署步驟：
 * 1. 在這份 Google Sheet 內，點選「擴充功能」→「Apps Script」
 * 2. 刪除預設內容，貼上這段程式碼
 * 3. 點選「部署」→「新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我（你的帳號）
 *    - 誰可以存取：任何人
 * 4. 部署後複製「網頁應用程式網址」，貼到 index.html 裡的 GAS_URL
 */
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

  return ContentService
    .createTextOutput(JSON.stringify({ result: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

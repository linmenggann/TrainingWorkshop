# Google Sheets 報名串接設定

目標試算表：`1v0_9-XG5DikOT4-Kk1mUToUXATkleCutJi4dwB3HMBA`
分頁名稱：`工作坊報名資料`

## 表頭（A1:I1）

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| 送出時間 | 姓名 | 機構名 | 職稱 | 負責的職類 | 是否擔任教學訓練計畫主持人 | 參與方式 | Email | 聯繫電話 |

表頭已寫入指定分頁。`Code.gs` 也會在收到第一筆資料時再次檢查表頭。

## 部署 Apps Script

1. 在 Google Sheets 點選「擴充功能」→「Apps Script」。
2. 將 `google-apps-script/Code.gs` 的完整內容貼入編輯器並儲存。
3. 點選右上角「部署」→「新增部署作業」。
4. 類型選擇「網頁應用程式」。
5. 「執行身分」選擇「我」，「誰可以存取」選擇「所有人」。
6. 完成授權與部署，複製結尾為 `/exec` 的網頁應用程式網址。
7. 開啟 `index.html`，將 `PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE` 替換成該 `/exec` 網址。
8. 再次提交並推送 `index.html`，即可正式收件。

## 驗證

- 在瀏覽器開啟 `/exec` 網址，應顯示「工作坊報名 API 運作中」。
- 從網頁送出一筆測試資料，確認資料出現在「工作坊報名資料」分頁第二列。
- 若更新 `Code.gs`，請在「管理部署作業」中建立新版本後重新部署；原 `/exec` 網址可維持不變。

## 儀表板資料 API

`dashboard.html` 會以 JSONP 讀取 Apps Script 的彙總資料，不會傳回姓名、Email 或電話。新增儀表板後，請將最新版 `Code.gs` 貼回 Apps Script，並在「管理部署作業」中建立新版本後重新部署。

重新部署後，可在 `/exec` 網址後加上 `?action=dashboard` 測試；成功時會回傳報名總數、參與方式、主持人狀態、職類與機構統計。

## 報名限額

活動限額設定為 4 名。網頁會透過 `?action=status` 查詢即時人數，達上限後顯示「額滿」並停用報名表單；`Code.gs` 也會在寫入資料的鎖定區段再次檢查人數，防止同時送出造成超額。

修改限額時，請同步調整 `Code.gs` 與 `index.html` 中的 `REGISTRATION_CAPACITY`，接著建立 Apps Script 新版本並重新部署。

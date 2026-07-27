# shekinah611 與 nankanchurch 付款後處理修復需求

## 現象

- 奉獻者完成付款後，對應收費單沒有更新。
- 同一筆付款也沒有發送 LINE 通知給奉獻者。
- 問題目前已知發生在租戶／參數 `shekinah611` 與 `nankanchurch`。
- 兩者的永豐豐收款「單筆」與「定期定額」付款皆有問題，需分別覆蓋 `QPayFeeProcessor` 與 `QPayDedicationBookingProcessor`，並優先檢查分流前的共用路徑。
- 本次範圍僅限永豐「豐收款／QPay」回呼；`MyPay` 是獨立金流模組，不在本次範圍內。

## 預期行為

- 成功付款回傳經驗證後，系統應更新正確的收費單付款狀態與付款資訊。
- 收費單更新成功後，系統應沿用既有流程發送 LINE 通知。
- 修復不得改變其他租戶或失敗付款的既有行為。

## 驗收條件

- 有自動化回歸測試能在修復前穩定重現問題並失敗。
- 修復後該測試通過，且既有測試／建置通過。
- 檢查付款回傳驗證、租戶路由、資料庫更新與 LINE 通知的完整控制流。
- 若僅靠程式碼與測試無法定位，新增最小且不洩漏敏感資訊的 trace 點，供實際付款重現時蒐證。

## 已確認證據

- 使用者提供的雲端 `app.config` 與目前工作樹 `QPayBackend/app.config` 完全相同；`611_XKeyID`、`NANKAN_XKeyID` 與 `SINOPAC_SITE` 均存在、無前後空白或控制字元。
- IIS 的 `web.config` 啟動 `QPayBackend.exe`，且程式使用 `ConfigurationManager.AppSettings`、沒有設定檔覆寫，因此執行階段讀取的是同目錄 `QPayBackend.exe.config`，不是 `app.config`。
- 本機既有 `bin/Output/QPayBackend.exe.config` 與建置前的 Debug runtime config 仍是舊 XKey；重新建置後，Debug runtime config 立即同步為新 XKey。
- 目前最強根因假說為：雲端只更新 `app.config`，但實際 runtime `QPayBackend.exe.config` 仍保留舊 XKey，或更新後尚未回收 IIS 程序。這會讓兩租戶在共同的永豐 Nonce／訂單查詢階段中斷，因此單筆與定期定額都不會進入 CRM 更新及 LINE 通知。

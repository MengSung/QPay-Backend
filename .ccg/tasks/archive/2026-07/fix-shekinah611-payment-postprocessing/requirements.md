# QPay 付款後處理失敗與發布設定混淆

## 問題

- `shekinah611` 與 `nankanchurch` 的單筆及定期定額付款，曾發生付款成功後未更新收費單、也未發送 LINE 通知。
- 程式實際讀取執行檔旁的 `QPayBackend.exe.config`，但維運時只更新了容易被誤認為有效設定的 `app.config`。
- 雲端主機的 `QPayBackend.exe.config` 因而保留舊 XKey，導致後續 QPay 查詢或驗證失敗。

## 核准範圍

- 僅修正永豐 QPay 的發布與設定驗證流程。
- 不修改付款、CRM、LINE 或 MyPay 的業務邏輯。
- 不在輸出、測試、文件或審查紀錄中顯示任何設定值、XKey 或完整 XML。

## 必要行為

- 每次 Release 發布先 clean，並使用全新的唯一輸出目錄。
- `app.config` 只作為建置來源，不得出現在部署目錄或 ZIP。
- 發布後逐一、精確比對來源與 `QPayBackend.exe.config` 的全部 `appSettings`。
- 任一 key 缺少、多出、值不同、必要設定為空或含外部空白時，必須回傳非 0 並中止發布，不能只顯示警告。
- 錯誤訊息只能包含固定分類與 key 名稱，不得包含 value。
- 驗證成功後才可建立 manifest、ZIP 與 SHA256；途中失敗不得留下可被誤認為完成的最終 ZIP。

## 驗收結果

- 負向整合測試已證明 stale runtime config 會讓完整 `dotnet publish` 以非 0 結束。
- 官方發布流程已建立經驗證的部署目錄、ZIP、manifest 與 SHA256。
- 部署目錄與 ZIP 均無 `app.config`，有效設定檔只有 `QPayBackend.exe.config`。

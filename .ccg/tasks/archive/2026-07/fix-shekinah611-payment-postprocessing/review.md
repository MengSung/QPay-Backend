# QPay 發布設定防呆審查

## 結論

- Critical：0。
- 已完成「乾淨發布、自動比對、移除部署包中的 `app.config`、驗證失敗回傳非 0」全部要求。
- 可以使用官方發布腳本產生的 ZIP 部署；不得把失敗的 `dotnet publish` 輸出目錄視為可部署成品。

## 驗證證據

- Windows PowerShell 5.1 parser：5 個腳本全部通過，0 個語法錯誤。
- TDD mutation 紅燈：暫時停用 `value-mismatch` 記錄後，負向整合測試因「stale runtime config 未中止 publish」正確失敗；隨即恢復原始檢查。
- 負向整合綠燈：注入假的 stale `611_XKeyID` 後，完整 `dotnet publish` 回傳非 0，只回報 key 與固定分類，未顯示 sentinel 或實際值。
- 全部 PowerShell 測試：3 個測試腳本通過；驗證器 7 個案例為 `7 passed, 0 failed`。
- 完整建置：Visual Studio 18 MSBuild 執行 `QPayBackend.sln` Release build，exit code 0。
- 正式發布：`Publish-QPayBackend.ps1 -Configuration Release` exit code 0，建立唯一的新目錄、manifest、ZIP 與 SHA256。
- 獨立成品核對：17 個 `appSettings` 完全一致；部署目錄與 ZIP 都沒有 `app.config`；ZIP 含 `QPayBackend.exe.config` 與 manifest；兩層 SHA256 均吻合；沒有 `.partial` 殘留。
- 範圍與秘密資料：11 個變更文字檔未包含來源敏感設定值；沒有 MyPay 範圍檔案變更；`git diff --check` 通過。

## 審查來源

- Gemini 外部審查完成，未回報 Critical。
- Claude 外部審查逾時；依使用者明確授權停止，不再延長雙模型等待。
- 兩輪獨立 `ccg-review`／本地審查均未發現 Critical。

## 已處理的審查意見

- ZIP 與 checksum 改用 `.partial` 暫存名稱；完成雜湊與 checksum 後才原子改名，最終 ZIP 最後才出現。
- 發布失敗會清除發布目錄、暫存 ZIP、暫存 checksum 與任何已建立的最終 artifact；清理例外不會取代固定的非 0 發布失敗結果。
- 新增完整 MSBuild 負向整合測試，證明驗證器非 0 確實向上傳播並中止 `dotnet publish`。

## 已知但接受的殘餘風險

- 直接執行 `dotnet publish --output ...` 時，MSBuild 會以非 0 中止，但不會自動回滾已寫入的輸出目錄。正式部署必須使用 `DotNetPublish-Release.bat`／`Publish-QPayBackend.ps1`，並嚴格禁止部署任何失敗命令的輸出。
- 若作業系統因檔案鎖定、ACL 或防毒軟體阻止刪除，失敗清理只能盡力執行；腳本仍會回傳非 0。最終 ZIP 只有在設定驗證、manifest、ZIP 雜湊與 checksum 全部成功後才以最後一步出現，降低誤部署風險。
- 專案仍有既有的 Newtonsoft.Json 版本／弱點警告、未 await 呼叫、隱藏 `Dispose`、舊 framework／ruleset 等警告；本次發布防呆沒有新增這些問題。

## 已驗證成品

- 部署目錄：`artifacts/QPayBackend-Release-20260727-124806-2ac45b30`
- ZIP：`artifacts/QPayBackend-Release-20260727-124806-2ac45b30.zip`
- SHA256：`artifacts/QPayBackend-Release-20260727-124806-2ac45b30.zip.sha256`

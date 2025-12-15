🎪 陽明國中 2年23班 園遊會點餐系統 (School Fair Order System)

這是一個基於 Google Apps Script (GAS) 開發的無伺服器 (Serverless) 點餐系統，專為園遊會場景設計。整合了 Google Sheets 作為後端資料庫，提供即時庫存管理、身分驗證及後台管理功能。

⚠️ 注意：本專案為 Google Apps Script 應用程式，無法透過 GitHub Pages 執行。GitHub 僅作為程式碼託管與備份使用。

🔗 專案連結

正式站點 (Web App): https://script.google.com/macros/s/AKfycbyIddAacnBAuXut1bohN62c00xhuQiT_c2_QPdYPwmT3yj8KzQmUjzkesbnF-yIC23hew/exec

管理後台:https://script.google.com/macros/s/AKfycbyIddAacnBAuXut1bohN62c00xhuQiT_c2_QPdYPwmT3yj8KzQmUjzkesbnF-yIC23hew/exec?page=admin

✨ 功能特色

📱 前台 (學生/顧客端)

RWD 響應式設計：使用 Tailwind CSS，完美支援手機與電腦版面。

即時庫存顯示：自動讀取試算表庫存，售完商品自動變灰並鎖定。

購物車系統：即時計算總金額，支援加減數量。

身分三重驗證：

針對校內學生進行「年級、班級、座號、姓名、學號」與「性別」的邏輯比對。

防止惡意下單或輸入錯誤資訊。

活動整合：嵌入 YouTube 宣傳影片與「神之手：匙來運轉」遊戲規則說明。

🛠 後台 (管理員端)

權限控制：僅限特定 Google 帳號存取後台。

訂單管理：查看所有訂單狀態，並可更新「取貨狀態」。

庫存管理：即時調整商品庫存數量。

防併發機制：使用 LockService 防止多人同時下單造成的庫存超賣問題。

📂 檔案結構

| 檔案名稱 | 說明 |
| Code.gs | 後端核心。負責處理 API 請求 (doGet)、與 Google Sheets 互動、驗證邏輯與庫存扣除。 |
| index.html | 前台介面。包含商品卡片渲染、購物車邏輯與使用者表單。 |
| admin.html | 後台介面。提供管理員查看訂單與修改庫存的儀表板。 |

🚀 如何部署

建立 Google Sheet：

建立一個新的 Google Sheet。

建立以下工作表 (Tabs)：

商品庫存 (欄位: ID, 名稱, 價格, 庫存, 圖片URL, 分類)

訂單紀錄 (系統自動寫入)

商品彙總 (系統自動統計)

國一名冊, 國二名冊, 國三名冊 (用於身分驗證)

建立 Google Apps Script：

在 Sheet 中點擊 擴充功能 > Apps Script。

將本專案的 Code.gs, index.html, admin.html 複製進去。

設定專案屬性 (Script Properties)：

為了安全起見，本專案不將敏感資料寫在程式碼中。

請至 GAS 編輯器 > 專案設定 (Project Settings) > 指令碼屬性 (Script Properties)。

新增以下屬性：

SPREADSHEET_ID: 您的試算表 ID。

ADMIN_EMAILS: 管理員 Email (若有多個請用逗號分隔)。

發布為 Web App (這是唯一的執行方式)：

點擊右上角 部署 > 新增部署。

類型選擇 網頁應用程式。

執行身分：我 (Me)。

誰可以存取：所有人 (Anyone)。

複製產生的網址，這就是您的網站連結。

⚠️ 注意事項

本專案使用 Google LockService 來處理併發請求，但在極高流量下 (如全校同時搶購) 仍建議進行壓力測試。

請勿將含有真實學生個資的 Google Sheet 權限設為「公開」。

🛠 技術棧

Backend: Google Apps Script (JavaScript)

Database: Google Sheets

Frontend: HTML5, Tailwind CSS via CDN

Icons: Heroicons (SVG)

Created for Yangming Junior High School, Class 223 Fair.

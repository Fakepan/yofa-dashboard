# 佑發空調-專案結案儀表板 | Claude Code 開發指南

## 🚀 快速開始

### 1. 在 Claude Code 中打開專案

```bash
# 在 Claude Code 終端機執行：
cd "你的專案路徑/yofa-dashboard-dev"

# 或者用 VS Code 打開
code .
```

### 2. 在瀏覽器中預覽

**方式 A：直接打開檔案**
```bash
# Windows
start index.html

# Mac
open index.html

# Linux
xdg-open index.html
```

**方式 B：用 Live Server（推薦）**
```bash
# 全域安裝 live-server
npm install -g live-server

# 在專案資料夾執行
live-server

# 自動打開 http://localhost:8080
```

**方式 C：用 Python HTTP 伺服器**
```bash
# Python 3
python -m http.server 8000

# 瀏覽器打開：http://localhost:8000
```

---

## 📁 專案結構

```
yofa-dashboard-dev/
├── index.html          ← 主應用（HTML + CSS + JS 一體）
├── ROADMAP.md          ← 開發規劃和優化清單
├── README.md           ← 本檔案
└── （未來新增）
    ├── src/
    │   ├── app.js      ← 核心邏輯（當拆分時）
    │   └── styles.css  ← 樣式（當拆分時）
    └── data/
        └── sample-data.json ← 樣本資料
```

---

## 💻 常見開發指令

### 查看檔案結構
```bash
# 列出所有檔案
ls -la

# 查看 HTML 檔案行數
wc -l index.html

# 搜尋特定程式碼
grep -n "function renderTable" index.html
```

### 編輯檔案
```bash
# 用 VS Code 編輯
code index.html

# 或用文字編輯器
nano index.html
```

### 分割檔案（如需要）
```bash
# 把 HTML 拆成多個檔案
# 1. 提取 CSS 到 style.css
# 2. 提取 JS 到 app.js
# 3. 建立 index.html 引入它們
```

---

## 🎯 常見任務

### 任務 1：新增一個功能
**例如：新增「搜尋」功能**

1. 在 Claude Code 說：
```
幫我在儀表板新增一個搜尋欄位，可以按案號或客戶名稱搜尋案件
```

2. Claude 會：
   - 修改 HTML（新增搜尋框）
   - 新增 JS 函數（搜尋邏輯）
   - 更新表格渲染（篩選結果）

3. 你在 Claude Code 中看到修改建議，按「確認」

### 任務 2：修正 Bug
**例如：表格有時候不顯示某些數據**

1. 說：
```
我發現費用計算不對，議價金額為 null 時應該用報價金額，請檢查 revenue() 函數
```

2. Claude 會找到相關程式碼並修正

### 任務 3：優化性能
1. 說：
```
表格有 100+ 筆資料時很慢，幫我加虛擬滾動或分頁
```

---

## 🧪 測試

### 手動測試清單

```
□ 新增案件 → 資料正確存入
□ 編輯成本 → 淨利自動計算
□ 刪除案件 → 圖表正確更新
□ 匯出 JSON → 能正確匯入
□ 匯入 JSON → 舊資料被覆蓋
□ 切換分頁 → 圖表正確顯示
□ 手機瀏覽 → 響應式正常
□ 關閉重開 → localStorage 正確恢復
```

### 瀏覽器 Console 檢查
```javascript
// 打開 F12 → Console，執行：

// 檢查目前資料
console.log(JSON.stringify(projects, null, 2))

// 檢查儲存
console.log(localStorage.getItem('yofa_dashboard_v1'))

// 測試計算
console.log(projects[0].revenue, projects[0].profit)
```

---

## 📊 開發流程建議

### 每次修改的步驟

```
1. 在 Claude Code 說出你的需求
   ↓
2. Claude 給出修改方案
   ↓
3. 你確認或提出調整
   ↓
4. Claude 應用修改到 index.html
   ↓
5. 你在瀏覽器重新整理查看效果
   ↓
6. 若 OK，繼續下一個功能；若不行，回報問題
```

---

## 🔧 Git 版本控制（可選）

如果你想追蹤每次修改：

```bash
# 初始化 Git
git init

# 加入所有檔案
git add .

# 第一次提交
git commit -m "Initial commit: v1.0 佑發空調儀表板"

# 每次大改之後提交
git commit -m "Add search feature"
git commit -m "Fix cost calculation bug"

# 查看歷史
git log --oneline
```

---

## 💾 備份方案

### 定期備份
```bash
# 複製到 Google Drive 同步資料夾
cp index.html "G:/My Drive/佑發空調系統/儀表板-開發中.html"

# 或上傳到 GitHub
git push origin main
```

---

## 📞 向 Claude 提出需求的格式

為了讓 Claude 更準確地幫你，建議這樣說：

**好的格式：**
> "在表格上方新增一個搜尋框，可以按案號或客戶名稱即時篩選，搜尋結果要同步到圖表"

**可以但不夠清楚：**
> "加個搜尋功能"

**超棒的格式：**
> "我想新增一個『獲利分析』頁面，顯示：
> 1. 各客戶的總淨利排名
> 2. 各月的毛利率趨勢
> 3. 虧損案件的警告標記"

---

## 🎓 學習資源

### 如果你想自己修改程式碼

- **HTML 基礎**：https://www.w3schools.com/html/
- **CSS 基礎**：https://www.w3schools.com/css/
- **JavaScript**：https://www.w3schools.com/js/
- **localStorage**：https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage

### 佑發專案特定的東西

- **佑發資料結構**：見 ROADMAP.md 中的「已實現功能」
- **計算邏輯**：在 `index.html` 中搜尋 `function revenue()`
- **資料儲存**：搜尋 `function saveLocal()`

---

## ❓ FAQ

**Q: 我修改了 index.html，但瀏覽器沒變化？**  
A: 按 Ctrl+Shift+R（強制重新整理）或清除快取

**Q: 資料都不見了？**  
A: localStorage 被清除了。用「📂 匯入資料」載入備份的 JSON

**Q: 能不能讓多台電腦同步？**  
A: 用「💾 匯出資料」+ Google Drive，或考慮升級到線上版本

**Q: 能加複雜的圖表嗎？**  
A: 可以，需要引入 Chart.js 或 Recharts 庫

---

## 🎉 下一步

1. **打開 index.html** 在瀏覽器中測試
2. **閱讀 ROADMAP.md** 看看想加什麼功能
3. **跟 Claude 說** 你的需求
4. **享受開發！** 🚀

---

**現在就開始吧！** 在 Claude Code 中說出你想要的功能，我們一起把它做出來！

---

*更新日期：2026-06-17*  
*維護者：佑發空調工程服務有限公司*

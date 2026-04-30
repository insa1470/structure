# 股權圖整併審核台

## 目前完成
- 前端單頁審核介面：`/webapp`
- 後端 API：`server.py`
- 任務資料持久化：`app_data/tasks/<task_id>/task.json`
- 使用者流程：
  - 上傳圖一、圖二
  - 建立任務
  - 查看總覽
  - 在網站上處理待確認與新增候選
  - 匯出審核結果 Excel

## 啟動方式
在專案根目錄執行：

```bash
python3 server.py
```

啟動後進入：

```text
http://127.0.0.1:8765
```

## 主要檔案
- `server.py`
  - 提供靜態頁面與 API
  - 接收圖片上傳
  - 建立任務與保存人工決策
- `webapp/index.html`
  - 使用者介面骨架
- `webapp/styles.css`
  - 介面樣式
- `webapp/app.js`
  - 前端互動、API 呼叫、Excel 匯出

## API
### `POST /api/tasks/analyze`
建立任務並接收兩張圖片。

目前分析模式先使用示範整併資料種子：
- 目的：先把使用流程打通
- 未來：在這個端點內接入真正 OCR 與整併流程

### `GET /api/demo-task`
建立一個示範任務，方便快速看流程。

### `GET /api/tasks/<task_id>`
取得任務完整資料。

### `POST /api/review-decision`
寫入待確認頁的人工決策。

### `POST /api/candidate-decision`
寫入新增候選頁的人工決策。

## 第二階段預留
第二階段目標不是現在完成，但資料結構已預留：

- `task.graph.nodes`
- `task.graph.edges`
- `task.graph.stage2`

未來第二階段可以直接基於：
- `master_rows`
- `review_decisions`
- `candidate_decisions`
- `graph.nodes`
- `graph.edges`

生成最終股權架構圖。

建議第二階段新增的 API：

### `POST /api/tasks/<task_id>/generate-chart`
用途：
- 讀取已審核完成的主表與圖結構
- 產生股權架構圖
- 回存輸出檔路徑與版本資訊

建議輸出：
- SVG
- PNG
- PowerPoint 用圖

## 目前限制
- 真正 OCR 辨識尚未接入；目前先用示範資料跑完整互動流程
- 上傳後的分析結果尚未依不同圖片內容改變
- 尚未加入登入、多人協作、版本追蹤

## 下一步建議
1. 把真正 OCR 與整併規則接到 `POST /api/tasks/analyze`
2. 把人工決策回寫到 `master_rows` 派生結果
3. 接上第二階段股權架構圖生成

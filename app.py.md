from flask import Flask, request, render\_template\_string

import openpyxl

import time

import uuid



app = Flask(\_\_name\_\_)



\# 短暫紀錄資料庫 (儲存於記憶體，重啟伺服器後清空)

history\_db = {}



HTML\_TEMPLATE = """

<!DOCTYPE html>

<html lang="zh-TW">

<head>

&#x20;   <meta charset="UTF-8">

&#x20;   <title>農林作物調查估價表 - 視覺化智能校對系統</title>

&#x20;   <style>

&#x20;       body { font-family: '微軟正黑體', sans-serif; background-color: #f4f7f6; margin: 0; padding: 20px; color: #333; }

&#x20;       .header-title { text-align: center; color: #2c3e50; margin-bottom: 20px; }

&#x20;       

&#x20;       /\* 版面配置：左側主畫面，右側紀錄面板 \*/

&#x20;       .layout-container { display: flex; gap: 20px; max-width: 1300px; margin: auto; align-items: flex-start; }

&#x20;       .main-content { flex: 1; background: #fff; padding: 30px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }

&#x20;       

&#x20;       .upload-box { border: 2px dashed #95a5a6; padding: 30px; text-align: center; border-radius: 10px; margin-bottom: 20px; background: #fafbfc; }

&#x20;       .success { color: #27ae60; font-weight: bold; background: #e9f7ef; padding: 10px; border-radius: 5px; }

&#x20;       .error-list { color: #c0392b; background: #fdf2f0; padding: 15px; border-left: 5px solid #e74c3c; border-radius: 4px; line-height: 1.6; }

&#x20;       .file-info { font-size: 16px; font-weight: bold; color: #34495e; padding: 10px 0; border-bottom: 2px solid #ecf0f1; margin-bottom: 20px; }



&#x20;       /\* 右側瀏覽紀錄面板 \*/

&#x20;       .history-panel { width: 320px; background: #fff; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); overflow: hidden; position: sticky; top: 20px; }

&#x20;       .history-header { background: #3498db; color: #fff; padding: 15px; font-size: 16px; font-weight: bold; display: flex; justify-content: space-between; cursor: pointer; }

&#x20;       .history-content { max-height: 70vh; overflow-y: auto; background: #fdfdfd; }

&#x20;       .history-item { padding: 12px 15px; border-bottom: 1px solid #eee; cursor: pointer; display: flex; flex-direction: column; transition: background 0.2s; }

&#x20;       .history-item:hover { background: #f0f8ff; }

&#x20;       .history-item.active { background: #e1f0fa; border-left: 4px solid #2980b9; }

&#x20;       .history-time { font-size: 12px; color: #7f8c8d; margin-top: 4px; }

&#x20;       

&#x20;       /\* 視覺化圖面專屬 CSS \*/

&#x20;       .map-container { overflow-x: auto; margin-top: 20px; border-radius: 8px; border: 1px solid #ddd; }

&#x20;       .excel-map { border-collapse: collapse; width: 100%; background: #fff; }

&#x20;       .excel-map th { background-color: #ecf0f1; border: 1px solid #bdc3c7; font-size: 12px; padding: 6px; color: #555; position: sticky; top: 0; z-index: 2; }

&#x20;       .excel-map td { border: 1px solid #ecf0f1; height: 30px; min-width: 50px; text-align: center; position: relative; font-size: 12px; }

&#x20;       

&#x20;       .cell-valid { background-color: #eafaf1; color: #27ae60; }

&#x20;       .cell-empty { background-color: #f9f9f9; color: #ccc; }

&#x20;       .cell-error { background-color: #fadbd8; color: #c0392b; border: 2px solid #e74c3c !important; font-weight: bold; cursor: help; }

&#x20;       

&#x20;       /\* Tooltip 特效 \*/

&#x20;       .tooltip {

&#x20;           position: absolute; background-color: rgba(0,0,0,0.85); color: #fff; padding: 8px 12px;

&#x20;           border-radius: 6px; font-size: 12px; font-weight: normal; bottom: 120%; left: 50%;

&#x20;           transform: translateX(-50%); white-space: nowrap; opacity: 0; visibility: hidden;

&#x20;           transition: opacity 0.2s ease; pointer-events: none; z-index: 9999;

&#x20;       }

&#x20;       .cell-error:hover .tooltip { opacity: 1; visibility: visible; }

&#x20;       .cell-error:hover { box-shadow: 2px 2px 10px rgba(231,76,60,0.4); z-index: 10; }

&#x20;   </style>

&#x20;   <script>

&#x20;       function toggleHistory() {

&#x20;           var content = document.getElementById('historyContent');

&#x20;           content.style.display = (content.style.display === 'none') ? 'block' : 'none';

&#x20;       }

&#x20;   </script>

</head>

<body>

&#x20;   <h2 class="header-title">📊 農林作物調查估價表 - 智能視覺化校對系統</h2>

&#x20;   

&#x20;   <div class="layout-container">

&#x20;       <!-- 左側主畫面 -->

&#x20;       <div class="main-content">

&#x20;           <div class="upload-box">

&#x20;               <form method="POST" enctype="multipart/form-data" action="/">

&#x20;                   <input type="file" name="excel\_file" accept=".xlsx" required>

&#x20;                   <button type="submit" style="padding: 8px 20px; font-weight: bold; cursor: pointer;">上傳並開始校對</button>

&#x20;               </form>

&#x20;           </div>



&#x20;           {% if record %}

&#x20;               <div class="file-info">📄 本次校對檔案：{{ record.filename }}</div>

&#x20;               

&#x20;               {% if record.error\_list|length == 0 %}

&#x20;                   <p class="success">✅ 完美！字體大小與循環格線格式皆完全符合規定。(系統已自動忽略 S 欄之後的內容)</p>

&#x20;               {% else %}

&#x20;                   <p>💡 系統已為您重繪 Excel 佈局 (僅校對至 S 欄)，將滑鼠移至<span style="color:#c0392b; font-weight:bold;">紅色格子</span>可查看錯誤原因。</p>

&#x20;                   

&#x20;                   <div class="map-container">

&#x20;                       <table class="excel-map">

&#x20;                           <tr>

&#x20;                               <th></th>

&#x20;                               {% for c in range(1, record.max\_col + 1) %}

&#x20;                                   <th>{{ get\_col\_letter(c) }}</th>

&#x20;                               {% endfor %}

&#x20;                           </tr>

&#x20;                           {% for r in range(1, record.max\_row + 1) %}

&#x20;                           <tr>

&#x20;                               <th>{{ r }}</th>

&#x20;                               {% for c in range(1, record.max\_col + 1) %}

&#x20;                                   {% set cell = record.map\_data.get((r, c), {'status': 'empty', 'text': '', 'error': ''}) %}

&#x20;                                   <td class="cell-{{ cell.status }}">

&#x20;                                       {{ cell.text }}

&#x20;                                       {% if cell.status == 'error' %}

&#x20;                                           <div class="tooltip">{{ cell.error }}</div>

&#x20;                                       {% endif %}

&#x20;                                   </td>

&#x20;                               {% endfor %}

&#x20;                           </tr>

&#x20;                           {% endfor %}

&#x20;                       </table>

&#x20;                   </div>



&#x20;                   <div class="error-list" style="margin-top: 20px;">

&#x20;                       <h4>📋 詳細錯誤報告清單：</h4>

&#x20;                       <ul>

&#x20;                           {% for err in record.error\_list %}

&#x20;                               <li>{{ err }}</li>

&#x20;                           {% endfor %}

&#x20;                       </ul>

&#x20;                   </div>

&#x20;               {% endif %}

&#x20;           {% endif %}

&#x20;       </div>



&#x20;       <!-- 右側歷史紀錄面板 -->

&#x20;       <div class="history-panel">

&#x20;           <div class="history-header" onclick="toggleHistory()">

&#x20;               <span>📂 瀏覽紀錄</span>

&#x20;               <span>▼</span>

&#x20;           </div>

&#x20;           <div class="history-content" id="historyContent">

&#x20;               {% for uid, hist in history.items()|reverse %}

&#x20;                   <div class="history-item {% if uid == current\_id %}active{% endif %}" onclick="window.location.href='/?id={{ uid }}'">

&#x20;                       <strong>{{ hist.filename }}</strong>

&#x20;                       <span class="history-time">{{ hist.time }}</span>

&#x20;                   </div>

&#x20;               {% else %}

&#x20;                   <div style="padding: 15px; text-align: center; color: #999; font-size: 13px;">尚無校對紀錄</div>

&#x20;               {% endfor %}

&#x20;           </div>

&#x20;       </div>

&#x20;   </div>

</body>

</html>

"""



@app.route("/", methods=\["GET", "POST"])

def index():

&#x20;   record\_id = request.args.get('id')

&#x20;   

&#x20;   if request.method == "POST":

&#x20;       file = request.files\["excel\_file"]

&#x20;       if file and file.filename.endswith('.xlsx'):

&#x20;           wb = openpyxl.load\_workbook(file, data\_only=False)

&#x20;           ws = wb.active

&#x20;           

&#x20;           error\_list = \[]

&#x20;           map\_data = {}

&#x20;           

&#x20;           # 【關鍵修正】強制限制最大校對欄位為 19 (即 S 欄)

&#x20;           max\_col = 19 

&#x20;           max\_row = min(ws.max\_row, 100) # 為了效能，圖面最多畫 100 列

&#x20;           merged\_ranges = ws.merged\_cells.ranges



&#x20;           for row in range(1, max\_row + 1):

&#x20;               # 這裡的迴圈範圍嚴格被鎖死在 1\~19，S 欄以後的資料 (T, U, V...) 完全不會被讀取與報錯

&#x20;               for col in range(1, max\_col + 1):

&#x20;                   cell = ws.cell(row=row, column=col)

&#x20;                   coord = cell.coordinate

&#x20;                   

&#x20;                   # 判斷合併儲存格與「有效底部列」

&#x20;                   is\_phantom = False

&#x20;                   eff\_bottom\_row = row

&#x20;                   

&#x20;                   for mr in merged\_ranges:

&#x20;                       min\_col, min\_row, max\_mr\_col, mr\_max\_row = mr.bounds

&#x20;                       # 若合併範圍超出 S 欄，我們也只在 A\~S 的範圍內處理

&#x20;                       if min\_row <= row <= mr\_max\_row and min\_col <= col <= max\_mr\_col:

&#x20;                           if row == min\_row and col == min\_col:

&#x20;                               is\_phantom = False

&#x20;                               eff\_bottom\_row = mr\_max\_row # 讓底線判斷以此格為準

&#x20;                           else:

&#x20;                               is\_phantom = True

&#x20;                           break



&#x20;                   if is\_phantom:

&#x20;                       map\_data\[(row, col)] = {'status': 'empty', 'text': ''}

&#x20;                       continue

&#x20;                       

&#x20;                   val = cell.value

&#x20;                   if val is None:

&#x20;                       map\_data\[(row, col)] = {'status': 'empty', 'text': ''}

&#x20;                       continue



&#x20;                   val\_str = str(val).strip()

&#x20;                   font = cell.font

&#x20;                   border = cell.border

&#x20;                   font\_size = font.size if font else None

&#x20;                   errors = \[]



&#x20;                   # 取得動態對照基準列 (處理大於等於 10 列的循環)

&#x20;                   ref\_row = 10 + ((row - 10) % 17) if row >= 10 else row

&#x20;                   eff\_ref\_row = 10 + ((eff\_bottom\_row - 10) % 17) if eff\_bottom\_row >= 10 else eff\_bottom\_row



&#x20;                   # ============== 【字體大小檢核】 ==============

&#x20;                   if col == 1 and row <= 9:

&#x20;                       pass # 條件 2：A1-A9無須檢核

&#x20;                   elif row <= 2 and col <= 17:

&#x20;                       if font\_size != 15:

&#x20;                           errors.append(f"標題字體應為 15pt (目前 {font\_size}pt)")

&#x20;                   elif row == 8 and col == 15:

&#x20;                       if font\_size != 10:

&#x20;                           errors.append(f"字體應為 10pt (目前 {font\_size}pt)")

&#x20;                   elif 21 <= ref\_row <= 25 and 3 <= col <= 19:

&#x20;                       if font\_size != 10:

&#x20;                           errors.append(f"字體應為 10pt (目前 {font\_size}pt)")

&#x20;                   else:

&#x20;                       pass # 其餘未定義區域預設放行



&#x20;                   # ============== 【格線檢核】 ==============

&#x20;                   # 條件 3：A21-S24、A26-S26 無底部格線 (以有效底部列 eff\_ref\_row 為準)

&#x20;                   is\_no\_bottom\_border = (21 <= eff\_ref\_row <= 24) or (eff\_ref\_row == 26)

&#x20;                   

&#x20;                   if is\_no\_bottom\_border:

&#x20;                       if border and border.bottom and border.bottom.style:

&#x20;                           errors.append(f"對照第 {eff\_ref\_row} 列，不應有底部格線")

&#x20;                   else:

&#x20;                       if row >= 10: # 一般明細資料列皆須底線

&#x20;                           if not border or not border.bottom or not border.bottom.style:

&#x20;                               errors.append(f"對照第 {eff\_ref\_row} 列，缺少底部格線")



&#x20;                   # 圖面預覽文字

&#x20;                   display\_text = val\_str\[:4] + '..' if len(val\_str) > 4 else val\_str

&#x20;                   

&#x20;                   if errors:

&#x20;                       err\_msg = "、".join(errors)

&#x20;                       map\_data\[(row, col)] = {'status': 'error', 'text': display\_text, 'error': err\_msg}

&#x20;                       error\_list.append(f"❌ {coord}: {err\_msg}")

&#x20;                   else:

&#x20;                       map\_data\[(row, col)] = {'status': 'valid', 'text': display\_text}



&#x20;           # 存入記憶體資料庫

&#x20;           uid = str(uuid.uuid4())\[:8]

&#x20;           history\_db\[uid] = {

&#x20;               'filename': file.filename,

&#x20;               'time': time.strftime("%H:%M:%S"),

&#x20;               'map\_data': map\_data,

&#x20;               'error\_list': error\_list,

&#x20;               'max\_row': max\_row,

&#x20;               'max\_col': max\_col

&#x20;           }

&#x20;           

&#x20;           return render\_template\_string(HTML\_TEMPLATE, record=history\_db\[uid], history=history\_db, current\_id=uid, get\_col\_letter=openpyxl.utils.get\_column\_letter)



&#x20;   # 點擊歷史紀錄切換

&#x20;   if record\_id and record\_id in history\_db:

&#x20;       return render\_template\_string(HTML\_TEMPLATE, record=history\_db\[record\_id], history=history\_db, current\_id=record\_id, get\_col\_letter=openpyxl.utils.get\_column\_letter)



&#x20;   # 首頁預設畫面

&#x20;   return render\_template\_string(HTML\_TEMPLATE, record=None, history=history\_db, current\_id=None)



if \_\_name\_\_ == "\_\_main\_\_":

&#x20;   app.run(debug=True, port=5000)


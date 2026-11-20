from flask import Flask, request, render_template_string, jsonify
import openpyxl
import time
import uuid
import math
import re
import gc
import ast
import operator
import datetime
import threading
import os
import tempfile

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 允許上傳最大 100 MB 檔案

# 資料庫 (全域變數)
history_db = {}
tasks_db = {}

# ==============================================================================
# 🏗️ 【系統檢核參數設定區】
# ==============================================================================
BUILDING_RULES = {
    "default_max_row": 85,          # 主估價表若未設定列印範圍，預設最多檢核列數
    "data_start_row": 11,           # 主估價表明細資料開始列數 (第 11 列起)
    # 主估價表視覺化顯示欄位：僅保留 B~L 與 N~Q (隱藏 A, M, R, S, T)
    "visible_cols": [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17],
    # 評點表視覺化顯示欄位：僅保留 D~M (D=4 ~ M=13)
    "rating_visible_cols": [4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
    "enable_col_width_check": True, # 開啟各欄欄寬檢核
    "col_width_xml_offset": 0.63,   # Excel 顯示欄寬與 XML 原始值之內距偏移量
    "col_width_tolerance": 0.25,    # 欄寬容許誤差
    # 自動跳過之工作表規則
    "skip_keywords": ["照片", "附表", "地籍資料", "值"],
    "skip_exact_names": {"A"},
}

# ==============================================================================
# 📏 【主估價表 - 預設各欄標準欄寬 (B~L, N~Q)】
# ==============================================================================
PRESET_COL_WIDTHS = {
    "B": 9.63,
    "C": 21.88,
    "D": 11.63,
    "E": 8.63,
    "F": 6.88,
    "G": 11.50,
    "H": 12.25,
    "I": 18.75,
    "J": 11.63,
    "K": 3.88,
    "L": 8.75,
    "N": 14.25,
    "O": 6.88,
    "P": 15.00,
    "Q": 15.25,
}

# ==============================================================================
# 📏 【建物評點表 - 預設各欄標準欄寬 (D~M)】
# ==============================================================================
RATING_PRESET_COL_WIDTHS = {
    "D": 3.75,
    "E": 9.25,
    "F": 15.88,
    "G": 5.75,
    "H": 15.88,
    "I": 5.75,
    "J": 15.88,
    "K": 5.75,
    "L": 15.88,
    "M": 5.75,
}

# ==============================================================================
# 📚 【建物項目單價與單位資料庫】
# 格式：(C欄「構造/規格」, D欄「用途」): {"unit": "K欄單位", "price": H欄單價}
# ==============================================================================
BUILDING_DB = {
    ("水泥地面", ""): {"unit": "㎡", "price": 480},
    # 後續您可自行往下擴充條目：
    # ("砌塊石圍牆(厚度35cm)", "擋土牆"): {"unit": "㎡", "price": 572},
    # ("鋼筋混凝土", "擋土牆"): {"unit": "㎡", "price": 6427},
    # ("深水井管徑10吋", ""): {"unit": "M", "price": 2310},
    # ("手動大門遷移費", ""): {"unit": "㎡", "price": 4100},
}

def custom_round(value, ndigits=0):
    """標準四捨五入 (避免 Python 銀行家捨入誤差)"""
    multiplier = 10 ** ndigits
    if value >= 0:
        rounded_val = math.floor(value * multiplier + 0.5 + 1e-12) / multiplier
    else:
        rounded_val = math.ceil(value * multiplier - 0.5 - 1e-12) / multiplier
    if ndigits == 0:
        return int(rounded_val)
    return rounded_val

def should_skip_sheet(sheet_name):
    """判斷是否跳過該工作表"""
    s = str(sheet_name or "").strip()
    if s in BUILDING_RULES["skip_exact_names"]:
        return True
    return any(kw in s for kw in BUILDING_RULES["skip_keywords"])

def is_rating_sheet(sheet_name):
    """判斷是否為建物評點表"""
    s = str(sheet_name or "").strip()
    return ("評點" in s)

def find_matching_rating_sheet(main_sheet_name, owner_name, rating_sheet_names):
    """
    尋找對應之建物評點表 (只要評點表名稱包含姓名即配對，無須完全一致)
    優先順序：1. 完全符合「{主表名}建物評點表」 -> 2. 包含主表名稱 -> 3. 包含 C3 地上物所有人姓名
    """
    main_clean = str(main_sheet_name or "").strip()
    owner_clean = str(owner_name or "").strip()

    exact_name = f"{main_clean}建物評點表"
    if exact_name in rating_sheet_names:
        return exact_name

    if main_clean:
        for r_name in rating_sheet_names:
            if main_clean in r_name:
                return r_name

    if owner_clean and len(owner_clean) >= 2:
        for r_name in rating_sheet_names:
            if owner_clean in r_name:
                return r_name

    return None

def get_cell_value_safe(ws, row, col, merged_ranges):
    """安全讀取儲存格內容（若位於合併儲存格內，自動抓取左上角主格的值）"""
    for mr in merged_ranges:
        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
            return ws.cell(row=min_row, column=min_col).value
    return ws.cell(row=row, column=col).value

def add_cell_error(all_calc_errors, row, col, msg, merged_ranges):
    """寫入儲存格錯誤（若目標格屬於合併儲存格的附屬格，自動導向左上角主格以免漏顯示）"""
    target_r, target_c = row, col
    for mr in merged_ranges:
        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
            target_r, target_c = min_row, min_col
            break
    if target_r in all_calc_errors and target_c in all_calc_errors[target_r]:
        all_calc_errors[target_r][target_c].append(msg)

def normalize_id_str(val):
    """正規化編號字串 (使 1、1.0、'1' 一致比對)"""
    if val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    s = str(val).strip()
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return s

def check_column_widths(ws, visible_cols, preset_widths):
    """
    讀取工作表各欄實際 XML 欄寬並展開連續區間，換算為 Excel 顯示欄寬後與指定標準值比對。
    """
    offset = BUILDING_RULES["col_width_xml_offset"]
    tol = BUILDING_RULES["col_width_tolerance"]
    enabled = BUILDING_RULES["enable_col_width_check"]

    default_xml_w = ws.sheet_format.defaultColWidth
    if default_xml_w is None:
        default_xml_w = 8.38 + offset
    else:
        default_xml_w = float(default_xml_w)

    raw_xml_widths = {c: default_xml_w for c in visible_cols}

    for key, dim in ws.column_dimensions.items():
        if dim.width is None:
            continue
        try:
            start_c = dim.min if dim.min is not None else openpyxl.utils.column_index_from_string(str(key))
            end_c = dim.max if dim.max is not None else start_c
        except Exception:
            continue
        for c_idx in range(start_c, end_c + 1):
            if c_idx in raw_xml_widths:
                raw_xml_widths[c_idx] = float(dim.width)

    col_width_status = {}
    col_width_error_list = []
    col_width_inspector = []

    for c_idx in visible_cols:
        col_letter = openpyxl.utils.get_column_letter(c_idx)
        raw_w = round(raw_xml_widths[c_idx], 2)
        est_display_w = round(max(0.0, raw_xml_widths[c_idx] - offset), 2)
        preset_w = preset_widths.get(col_letter)

        is_err = False
        err_msg = ""
        if enabled and preset_w is not None:
            if abs(est_display_w - preset_w) > tol:
                is_err = True
                err_msg = f"{col_letter}欄寬度不符：目前約 {est_display_w:.2f} (XML {raw_w:.2f})，預設應為 {preset_w:.2f}"
                col_width_error_list.append({
                    "col": col_letter,
                    "actual": f"{est_display_w:.2f}",
                    "raw_xml": f"{raw_w:.2f}",
                    "preset": f"{preset_w:.2f}",
                    "diff": f"{(est_display_w - preset_w):+.2f}",
                    "msg": err_msg
                })

        col_width_status[c_idx] = {
            "is_error": is_err,
            "est_display": f"{est_display_w:.2f}",
            "raw_xml": f"{raw_w:.2f}",
            "preset": f"{preset_w:.2f}" if preset_w is not None else "未設定",
            "error": err_msg
        }
        col_width_inspector.append({
            "col": col_letter,
            "est_display": f"{est_display_w:.2f}",
            "raw_xml": f"{raw_w:.2f}",
            "preset": f"{preset_w:.2f}" if preset_w is not None else "未設定",
            "is_error": is_err
        })

    return col_width_status, col_width_error_list, col_width_inspector

def is_pure_number(val):
    """判斷是否為純數值 (排除 None、bool、datetime、文字)"""
    if val is None or isinstance(val, (bool, datetime.date, datetime.datetime)):
        return False
    if isinstance(val, (int, float)):
        return not math.isnan(val)
    try:
        s = str(val).replace(',', '').strip()
        if s == "":
            return False
        float(s)
        return True
    except Exception:
        return False

def to_float(val):
    """將數值或千分位字串安全轉為 float"""
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        return float(val)
    s = str(val).replace(',', '').strip()
    return float(s)

def parse_rate_value(val):
    """解析 L 欄發放比例 (支援 0.7、1、1.0 或 '70%'、'100%')"""
    if val is None or isinstance(val, (bool, datetime.date, datetime.datetime)):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s == "":
        return None
    if s.endswith('%'):
        try:
            return float(s[:-1].strip()) / 100.0
        except Exception:
            return None
    try:
        return float(s.replace(',', ''))
    except Exception:
        return None

# 安全算術 AST 求值器
_ALLOWED_BINOPS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
}
_ALLOWED_UNARYOPS = {
    ast.UAdd: operator.pos,
    ast.USub: operator.neg,
}

def _safe_eval_node(node):
    if isinstance(node, ast.Expression):
        return _safe_eval_node(node.body)
    elif isinstance(node, ast.Constant) and isinstance(node.value, (int, float)) and not isinstance(node.value, bool):
        return float(node.value)
    elif isinstance(node, ast.BinOp) and type(node.op) in _ALLOWED_BINOPS:
        left = _safe_eval_node(node.left)
        right = _safe_eval_node(node.right)
        if isinstance(node.op, ast.Div) and abs(right) < 1e-12:
            raise ZeroDivisionError("除數不可為 0")
        return _ALLOWED_BINOPS[type(node.op)](left, right)
    elif isinstance(node, ast.UnaryOp) and type(node.op) in _ALLOWED_UNARYOPS:
        return _ALLOWED_UNARYOPS[type(node.op)](_safe_eval_node(node.operand))
    else:
        raise ValueError("包含不合法的運算結構")

def validate_and_eval_formula(raw_val):
    """檢核計算式是否合理並計算結果 (適用於主表 I 欄與評點表 F37/H37/J37/L37)"""
    if raw_val is None:
        return True, None, ""
    if isinstance(raw_val, (int, float)) and not isinstance(raw_val, bool):
        return True, float(raw_val), ""

    expr = str(raw_val).strip()
    if expr == "":
        return True, None, ""

    if expr.startswith('='):
        expr_body = expr[1:].strip()
        if re.search(r'[A-Za-z]', expr_body):
            return True, None, ""
        expr = expr_body

    compact = re.sub(r'\s+', '', expr)
    if compact == "":
        return True, None, ""

    fullwidth_match = re.findall(r'[×÷＋－＊／（）]', compact)
    if fullwidth_match:
        return False, None, f"計算式包含全形符號 ({''.join(set(fullwidth_match))})，請改為半形"

    invalid_chars = re.findall(r'[^0-9\.\+\-\*\/\(\)]', compact)
    if invalid_chars:
        return False, None, f"計算式包含非合法字元 ({''.join(set(invalid_chars))})"

    dup_op = re.search(r'[\+\-\*\/]{2,}', compact)
    if dup_op:
        return False, None, f"計算式不合理：出現連續運算符號 ({dup_op.group()})"

    if re.match(r'^[\+\*\/]', compact):
        return False, None, f"計算式不合理：開頭不可為運算符號 ({compact[0]})"
    if re.search(r'[\+\-\*\/]$', compact):
        return False, None, f"計算式不合理：運算符號 ({compact[-1]}) 後方缺少數字"

    depth = 0
    for ch in compact:
        if ch == '(':
            depth += 1
        elif ch == ')':
            depth -= 1
            if depth < 0:
                return False, None, "計算式不合理：括號順序錯誤或不成對"
    if depth != 0:
        return False, None, "計算式不合理：左右括號數量不成對"

    if '()' in compact:
        return False, None, "計算式不合理：出現空括號 ()"
    if re.search(r'\([\+\*\/]', compact):
        return False, None, "計算式不合理：左括號後方不可直接接運算符號"
    if re.search(r'[\+\-\*\/]\)', compact):
        return False, None, "計算式不合理：右括號前方缺少數字"
    if re.search(r'\d\(', compact) or re.search(r'\)\d', compact) or ')(' in compact:
        return False, None, "計算式不合理：括號與數字之間缺少乘號 (*)"

    if '..' in compact or re.search(r'\d+\.\d+\.', compact):
        return False, None, "計算式不合理：小數點格式錯誤"
    if re.search(r'(?<!\d)\.(?!\d)', compact):
        return False, None, "計算式不合理：出現孤立的小數點"

    try:
        tree = ast.parse(compact, mode='eval')
        val = _safe_eval_node(tree)
        return True, val, ""
    except ZeroDivisionError:
        return False, None, "計算式不合理：除數不可為 0"
    except Exception:
        return False, None, "計算式語法無法解析，請檢查公式格式"

def format_cell_display(val, col_idx, row_idx, sheet_type="main"):
    """美化視覺化網格中的儲存格文字顯示"""
    if val is None:
        return ""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%Y/%m/%d")
    if isinstance(val, float):
        if sheet_type == "main" and col_idx == 12 and row_idx >= 11 and 0 <= val <= 2:
            pct = custom_round(val * 100, 1)
            return f"{int(pct)}%" if pct == int(pct) else f"{pct}%"
        if val == int(val):
            s = f"{int(val):,}" if sheet_type == "main" and abs(val) >= 1000 and col_idx in (8, 14, 17) else str(int(val))
        else:
            s = f"{round(val, 4)}"
    elif isinstance(val, int) and not isinstance(val, bool):
        if sheet_type == "main" and col_idx == 12 and row_idx >= 11 and val in (0, 1):
            return f"{val * 100}%"
        s = f"{val:,}" if sheet_type == "main" and abs(val) >= 1000 and col_idx in (8, 14, 17) else str(val)
    else:
        s = str(val).replace('\n', ' ').strip()
    return s[:9] + '..' if len(s) > 9 else s

def extract_max_row_from_print_area(print_area):
    """解析單一或多段列印範圍取得最大列數"""
    if not print_area:
        return None
    if isinstance(print_area, (list, tuple)):
        areas = [str(a) for a in print_area]
    elif isinstance(print_area, str):
        areas = [a.strip() for a in print_area.split(',') if a.strip()]
    else:
        return None

    max_r = None
    for area in areas:
        part = area.split('!')[-1]
        end_coord = part.split(':')[1] if ':' in part else part
        m = re.search(r'\d+', end_coord)
        if m:
            r_num = int(m.group())
            if max_r is None or r_num > max_r:
                max_r = r_num
    return max_r

def sum_rating_block_scores(ws_r_data, score_col):
    """需求 2：加總評點表指定評點欄第 7 列至第 33 列之數值，回傳 (has_any_score, total_rounded_1)"""
    total = 0.0
    has_any = False
    for r in range(7, 34):
        val = ws_r_data.cell(row=r, column=score_col).value
        if is_pure_number(val):
            total += to_float(val)
            has_any = True
    return has_any, custom_round(total, 1)

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <title>建築改良物調查估價表 - 智能視覺化校對系統</title>
    <style>
        body { font-family: '微軟正黑體', sans-serif; background-color: #f4f7f6; margin: 0; padding: 14px 16px; color: #333; scroll-behavior: smooth; }
        .header-title { text-align: center; color: #2c3e50; margin: 4px 0 16px 0; font-size: 22px; }
        .layout-container { display: flex; gap: 16px; max-width: 98%; width: 100%; margin: auto; align-items: flex-start; }
        .main-content { flex: 1; background: #fff; padding: 22px 24px; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: relative; min-width: 0; }
        .upload-box { border: 2px dashed #3498db; padding: 22px; text-align: center; border-radius: 10px; margin-bottom: 18px; background: #f8fbff; transition: 0.3s; }
        
        /* 載入動畫 */
        #loadingOverlay { display: none; text-align: center; padding: 40px 20px; }
        .spinner { border: 6px solid #f3f3f3; border-top: 6px solid #3498db; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; margin: 0 auto 20px auto; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        .progress-bar-container { width: 80%; background-color: #e0e0e0; border-radius: 8px; margin: 15px auto; height: 20px; overflow: hidden; }
        .progress-bar { height: 100%; background-color: #2ecc71; width: 0%; transition: width 0.4s ease; }

        /* 需求 3：報表上方工具列與「一鍵只顯示有錯誤的工作表」按鈕 */
        .report-toolbar { display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 10px; padding-bottom: 12px; margin-bottom: 12px; border-bottom: 2px solid #ecf0f1; }
        .toolbar-controls { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
        .filter-status { font-size: 13px; color: #555; background: #f2f4f6; padding: 5px 10px; border-radius: 5px; }
        .btn-filter { padding: 8px 14px; border: 2px solid #e67e22; border-radius: 6px; background: #fff; color: #d35400; font-size: 13px; font-weight: bold; cursor: pointer; transition: 0.2s; font-family: inherit; }
        .btn-filter:hover { background: #fef5e7; }
        .btn-filter.active { background: #e74c3c; border-color: #c0392b; color: #fff; }
        .btn-filter.active:hover { background: #c0392b; }
        #noErrorSheetsBanner { display: none; background: #e9f7ef; color: #1e8449; border: 1px solid #a9dfbf; padding: 14px 18px; border-radius: 6px; font-weight: bold; margin-bottom: 16px; text-align: center; }

        .skipped-info { background: #f2f4f6; border-left: 4px solid #7f8c8d; padding: 8px 14px; border-radius: 5px; font-size: 12px; color: #555; margin-bottom: 16px; }
        .success { color: #27ae60; font-weight: bold; background: #e9f7ef; padding: 10px 14px; border-radius: 5px; }
        .special-warning { background-color: #f3e5f5 !important; color: #6a1b9a !important; border: 2px solid #9c27b0 !important; font-weight: bold; cursor: help; }
        
        .correction-panel { background: #fff3cd; border-left: 6px solid #ffc107; padding: 14px 18px; border-radius: 6px; margin-bottom: 18px; }
        .correction-panel h4 { margin-top: 0; margin-bottom: 10px; color: #856404; font-size: 15px; }
        .correction-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 4px; overflow: hidden; }
        .correction-table th, .correction-table td { padding: 8px 10px; border: 1px solid #ffeeba; text-align: center; font-size: 13px; }
        .correction-table th { background: #ffe8a1; color: #856404; }

        /* 欄寬異常專屬警告面板 */
        .width-panel { background: #fdecea; border-left: 6px solid #e74c3c; padding: 14px 18px; border-radius: 6px; margin-bottom: 18px; }
        .width-panel h4 { margin-top: 0; margin-bottom: 10px; color: #922b21; font-size: 15px; }
        .width-table { width: 100%; border-collapse: collapse; background: #fff; border-radius: 4px; overflow: hidden; }
        .width-table th, .width-table td { padding: 7px 10px; border: 1px solid #f5c6cb; text-align: center; font-size: 13px; }
        .width-table th { background: #f8d7da; color: #922b21; }
        
        /* 結構探測器面板 */
        .inspector-box { background: #eef6ff; border: 1px solid #b8daff; border-radius: 6px; padding: 10px 14px; margin-bottom: 16px; }
        .inspector-toggle { background: #2980b9; color: #fff; border: none; padding: 6px 12px; border-radius: 4px; cursor: pointer; font-size: 13px; font-weight: bold; }
        .inspector-toggle:hover { background: #1f6391; }
        .inspector-content { display: none; margin-top: 10px; max-height: 360px; overflow-y: auto; background: #fff; border: 1px solid #ccc; border-radius: 4px; padding: 10px; font-size: 12px; }
        .inspector-table { width: 100%; border-collapse: collapse; margin-bottom: 14px; }
        .inspector-table th, .inspector-table td { border: 1px solid #ddd; padding: 5px 8px; text-align: left; }
        .inspector-table th { background: #f1f5f9; position: sticky; top: 0; }

        /* 右側精簡瀏覽紀錄面板 */
        .history-panel { width: 215px; flex-shrink: 0; background: #fff; border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); position: sticky; top: 16px; overflow: hidden; }
        .panel-actions { display: flex; gap: 6px; padding: 8px; background: #ecf0f1; border-bottom: 1px solid #dcdde1; }
        .action-btn { flex: 1; padding: 6px 4px; border: none; border-radius: 5px; font-size: 12px; font-weight: bold; cursor: pointer; transition: 0.2s; font-family: inherit; white-space: nowrap; }
        .btn-top { background: #2ecc71; color: #fff; }
        .btn-top:hover { background: #27ae60; }
        .btn-clear { background: #e74c3c; color: #fff; }
        .btn-clear:hover { background: #c0392b; }
        
        .history-header { background: #2c3e50; color: #fff; padding: 11px 12px; font-size: 14px; font-weight: bold; display: flex; justify-content: space-between; cursor: pointer; }
        .history-content { max-height: 75vh; overflow-y: auto; background: #fdfdfd; }
        .history-item { padding: 10px 10px; border-bottom: 1px solid #eee; display: flex; flex-direction: column; transition: background 0.2s; font-size: 13px; word-break: break-all; }
        .history-item:hover { background: #f0f8ff; }
        .history-item.active { background: #e1f0fa; border-left: 4px solid #2980b9; }
        .history-time { font-size: 11px; color: #7f8c8d; }
        
        /* 視覺化 Excel 網格 */
        .map-container { overflow-x: auto; border-radius: 8px; border: 1px solid #ddd; margin-top: 12px; }
        .excel-map { border-collapse: collapse; width: 100%; table-layout: fixed; background: #fff; }
        .excel-map th { background-color: #ecf0f1; border: 1px solid #bdc3c7; font-size: 12px; padding: 6px 4px; position: sticky; top: 0; z-index: 2; }
        .excel-map th.row-num { width: 34px; }
        .excel-map th.col-width-error { background-color: #fadbd8 !important; color: #c0392b !important; border: 2px solid #e74c3c !important; cursor: help; }
        .excel-map th.col-width-error:hover { z-index: 25; }
        .excel-map th.col-width-error:hover .tooltip { opacity: 1; visibility: visible; }
        .excel-map th .tooltip { bottom: auto; top: 115%; }
        .excel-map th:nth-child(-n+3) .tooltip { left: 0; transform: none; }
        .excel-map th:nth-last-child(-n+3) .tooltip { left: auto; right: 0; transform: none; }

        .excel-map td { border: 1px solid #e5e8e8; height: 32px; padding: 2px 4px; text-align: center; position: relative; font-size: 12px; overflow: visible; word-break: break-word; }
        .cell-valid { background-color: #eafaf1; color: #27ae60; }
        .cell-empty { background-color: #f9f9f9; color: #ccc; }
        .cell-error { background-color: #fadbd8; color: #c0392b; border: 2px solid #e74c3c !important; font-weight: bold; cursor: help; }
        .tooltip { position: absolute; background-color: rgba(0,0,0,0.9); color: #fff; padding: 8px 12px; border-radius: 6px; font-size: 12px; font-weight: normal; bottom: 118%; left: 50%; transform: translateX(-50%); white-space: nowrap; opacity: 0; visibility: hidden; z-index: 9999; pointer-events: none; box-shadow: 0 2px 8px rgba(0,0,0,0.3); }
        .excel-map tr:nth-child(-n+6) .tooltip { bottom: auto; top: 118%; }
        .excel-map td:nth-child(-n+3) .tooltip { left: 0; transform: none; }
        .excel-map td:nth-last-child(-n+3) .tooltip { left: auto; right: 0; transform: none; }
        .cell-error:hover .tooltip, .special-warning:hover .tooltip { opacity: 1; visibility: visible; }
        
        .sheet-container { border: 2px solid #e0e0e0; border-radius: 8px; padding: 18px; margin-bottom: 30px; background: #fafbfc; }
        .sheet-container.sheet-rating { border-color: #a9cce3; background: #f7fbfe; }
        .sheet-title { margin-top: 0; color: #34495e; border-bottom: 2px solid #bdc3c7; padding: 10px; border-radius: 4px; transition: 0.2s; scroll-margin-top: 20px; cursor: pointer; display: flex; justify-content: space-between; font-size: 16px; }
        .sheet-title:hover { background-color: #e2e8f0; }
        .type-tag { display: inline-block; padding: 2px 8px; border-radius: 4px; font-size: 12px; margin-right: 6px; vertical-align: middle; }
        .tag-main { background: #d5f5e3; color: #1e8449; }
        .tag-rating { background: #d6eaf8; color: #1b4f72; }

        .meta-badges { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 12px; font-size: 12px; }
        .badge { background: #eef2f7; color: #34495e; padding: 4px 10px; border-radius: 12px; border: 1px solid #d5dde5; }
        .badge-ok { background: #e9f7ef; color: #1e8449; border-color: #a9dfbf; }
        .badge-err { background: #fdecea; color: #c0392b; border-color: #f5b7b1; font-weight: bold; }
    </style>
</head>
<body id="pageTop">
    <h2 class="header-title">🏢 建築改良物調查估價表＆評點表 - 智能視覺化校對系統</h2>
    <div class="layout-container">
        <div class="main-content">
            
            <!-- 上傳區塊 -->
            <div class="upload-box" id="uploadBox">
                <form id="uploadForm">
                    <input type="file" name="excel_file" id="excelFile" accept=".xlsx, .xlsm" required>
                    <button type="submit" style="padding: 8px 20px; font-weight: bold; cursor: pointer; background: #2980b9; color: #fff; border: none; border-radius: 5px;">上傳建物估價表 (.xlsx / .xlsm) 並開始校對</button>
                </form>
            </div>

            <!-- 載入動畫區塊 -->
            <div id="loadingOverlay">
                <div class="spinner"></div>
                <h3 style="color: #2980b9;">建物估價表與評點表正在背景火速解析與跨表核對中，請稍候... (<span id="progressText">0%</span>)</h3>
                <div class="progress-bar-container"><div class="progress-bar" id="progressBar"></div></div>
                <p style="color: #e74c3c; font-weight: bold; font-size: 11pt;">⚠️ 若工作表數量龐大，可能需要數分鐘處理，請勿關閉或重新整理網頁！</p>
                <p id="currentSheetText" style="color: #7f8c8d; font-size: 13px;"></p>
            </div>

            <script>
                document.getElementById('uploadForm').onsubmit = async function(e) {
                    e.preventDefault();
                    let formData = new FormData(this);
                    
                    document.getElementById('uploadBox').style.display = 'none';
                    document.getElementById('loadingOverlay').style.display = 'block';
                    if(document.getElementById('reportContainer')) {
                        document.getElementById('reportContainer').style.display = 'none';
                    }

                    try {
                        let response = await fetch('/upload', { method: 'POST', body: formData });
                        let data = await response.json();
                        
                        if (data.task_id) {
                            checkTaskStatus(data.task_id);
                        } else {
                            alert("上傳失敗: " + (data.error || "未知錯誤"));
                            location.reload();
                        }
                    } catch (err) {
                        alert("網路連線錯誤，請重試。");
                        location.reload();
                    }
                };

                async function checkTaskStatus(taskId) {
                    try {
                        let response = await fetch('/status/' + taskId);
                        let data = await response.json();
                        
                        if (data.status === 'processing') {
                            document.getElementById('progressText').innerText = data.progress + "%";
                            document.getElementById('progressBar').style.width = data.progress + "%";
                            if (data.current_sheet) {
                                document.getElementById('currentSheetText').innerText = "目前正在處理: " + data.current_sheet;
                            }
                            setTimeout(() => checkTaskStatus(taskId), 2000);
                        } else if (data.status === 'completed') {
                            document.getElementById('progressText').innerText = "100%";
                            document.getElementById('progressBar').style.width = "100%";
                            window.location.href = '/?id=' + data.record_id;
                        } else if (data.status === 'error') {
                            alert("背景運算發生錯誤: " + data.message);
                            location.reload();
                        }
                    } catch (err) {
                        setTimeout(() => checkTaskStatus(taskId), 2000);
                    }
                }

                function toggleHistory() {
                    let content = document.getElementById('historyContent');
                    content.style.display = (content.style.display === 'none') ? 'block' : 'none';
                }

                function toggleSheet(sheetId) {
                    let content = document.getElementById('content-' + sheetId);
                    let icon = document.getElementById('icon-' + sheetId);
                    if (content.style.display === 'none') {
                        content.style.display = 'block';
                        icon.innerHTML = '▼';
                    } else {
                        content.style.display = 'none';
                        icon.innerHTML = '▲';
                    }
                }

                function toggleInspector(sheetId) {
                    let box = document.getElementById('inspector-' + sheetId);
                    box.style.display = (box.style.display === 'block') ? 'none' : 'block';
                }

                function scrollToTop() {
                    const uploadBox = document.getElementById('uploadBox');
                    if (uploadBox) {
                        uploadBox.style.display = 'block';
                    }
                    window.scrollTo({ top: 0, behavior: 'smooth' });
                }

                async function confirmClearHistory() {
                    if (confirm("⚠️ 確定要清除所有瀏覽紀錄與暫存資料嗎？\\n\\n清除後可釋放伺服器記憶體避免卡頓，但已校對的歷史紀錄將無法復原。")) {
                        try {
                            let response = await fetch('/clear_history', { method: 'POST' });
                            let data = await response.json();
                            if (data.status === 'cleared') {
                                window.location.href = '/';
                            }
                        } catch (err) {
                            alert("清除紀錄時發生錯誤，請重試。");
                        }
                    }
                }

                // 需求 3：一鍵只顯示有錯誤的工作表（預設關閉，每次點擊皆跳出確認提醒）
                let onlyErrorsMode = false;
                function toggleErrorSheetsFilter() {
                    const sheets = document.querySelectorAll('.sheet-container');
                    let errorSheetCount = 0;
                    sheets.forEach(el => {
                        if (el.getAttribute('data-has-error') === 'true') {
                            errorSheetCount++;
                        }
                    });

                    if (!onlyErrorsMode) {
                        const confirmMsg = "⚠️ 確定要開啟「只顯示有錯誤的工作表」篩選嗎？\\n\\n本次共校對 " + sheets.length + " 張工作表，其中 " + errorSheetCount + " 張有異常。\\n開啟後將暫時隱藏所有無錯誤（✅）的工作表。";
                        if (!confirm(confirmMsg)) {
                            return;
                        }

                        onlyErrorsMode = true;
                        sheets.forEach(el => {
                            if (el.getAttribute('data-has-error') !== 'true') {
                                el.style.display = 'none';
                            } else {
                                el.style.display = 'block';
                            }
                        });

                        const btn = document.getElementById('errorFilterBtn');
                        btn.classList.add('active');
                        btn.innerText = "📋 恢復顯示全部工作表";
                        document.getElementById('filterStatusText').innerText = "篩選中：僅顯示有錯誤 (" + errorSheetCount + " / " + sheets.length + " 張)";

                        const emptyBanner = document.getElementById('noErrorSheetsBanner');
                        if (emptyBanner) {
                            emptyBanner.style.display = (errorSheetCount === 0) ? 'block' : 'none';
                        }
                    } else {
                        if (!confirm("⚠️ 確定要關閉篩選並「恢復顯示全部工作表」嗎？")) {
                            return;
                        }

                        onlyErrorsMode = false;
                        sheets.forEach(el => {
                            el.style.display = 'block';
                        });

                        const btn = document.getElementById('errorFilterBtn');
                        btn.classList.remove('active');
                        btn.innerText = "🔍 一鍵只顯示有錯誤的工作表";
                        document.getElementById('filterStatusText').innerText = "顯示中：全部工作表 (共 " + sheets.length + " 張)";

                        const emptyBanner = document.getElementById('noErrorSheetsBanner');
                        if (emptyBanner) {
                            emptyBanner.style.display = 'none';
                        }
                    }
                }
            </script>

            <!-- 渲染報表區塊 -->
            {% if record %}
            <div id="reportContainer">
                <div class="report-toolbar">
                    <div class="file-info" style="font-size:17px; font-weight: bold;">📄 本次校對完成：{{ record.filename }}</div>
                    <div class="toolbar-controls">
                        <span class="filter-status" id="filterStatusText">顯示中：全部工作表 (共 {{ record.sheets|length }} 張)</span>
                        <button type="button" id="errorFilterBtn" class="btn-filter" onclick="toggleErrorSheetsFilter()">🔍 一鍵只顯示有錯誤的工作表</button>
                    </div>
                </div>
                
                {% if record.skipped_sheets %}
                <div class="skipped-info">
                    ⏭️ 已自動略過 {{ record.skipped_sheets|length }} 張非校對工作表（照片 / 附表 / 地籍資料 / 值 / A）：{{ record.skipped_sheets|join('、') }}
                </div>
                {% endif %}

                <div id="noErrorSheetsBanner">
                    🎉 本次校對的所有工作表皆無異常，目前沒有需要顯示的錯誤工作表！
                </div>

                {% for sheet_name, s_data in record.sheets.items() %}
                {% set s_idx = loop.index %}
                {% set has_err = (s_data.error_list|length > 0) or (s_data.correction_list|length > 0) or (s_data.col_width_error_list|length > 0) %}
                <div class="sheet-container {% if s_data.sheet_type == 'rating' %}sheet-rating{% endif %}" data-has-error="{{ 'true' if has_err else 'false' }}">
                    <h3 class="sheet-title" id="sheet-{{ s_idx }}" onclick="toggleSheet('{{ s_idx }}')">
                        <span>
                            {% if s_data.sheet_type == 'rating' %}
                                <span class="type-tag tag-rating">建物評點表</span>
                            {% else %}
                                <span class="type-tag tag-main">建物估價表</span>
                            {% endif %}
                            {{ sheet_name }} (列印範圍：{{ s_data.print_area_str }})
                        </span>
                        <span id="icon-{{ s_idx }}">▼</span>
                    </h3>
                    <div id="content-{{ s_idx }}">
                        
                        <!-- 關鍵欄位判讀狀態摘要 -->
                        <div class="meta-badges">
                            {% if s_data.sheet_type == 'rating' %}
                                <span class="badge">檢核範圍：D～M 欄 (第 1～{{ s_data.max_row }} 列)</span>
                                <span class="badge">評點加總：第 7～33 列</span>
                                <span class="badge">字體檢核：D5:M35 須為 12pt</span>
                                <span class="badge">已填寫棟數：{{ s_data.meta.active_blocks }} 棟</span>
                            {% else %}
                                <span class="badge">H1：{{ s_data.meta.h1_text or '(空)' }}</span>
                                <span class="badge">S24：{{ s_data.meta.s24_text or '(空)' }}</span>
                                <span class="badge">Q5 面積：{{ s_data.meta.q5_text or '(空)' }}</span>
                                <span class="badge">S/T 合計：{{ s_data.meta.st_total_text or '未偵測' }}</span>
                                <span class="badge">對應評點表：{{ s_data.meta.matched_rating or '無須配對 / 未找到' }}</span>
                            {% endif %}
                            {% if s_data.col_width_error_list|length == 0 %}
                                <span class="badge badge-ok">📏 欄寬檢核：全數符合</span>
                            {% else %}
                                <span class="badge badge-err">📏 欄寬異常：{{ s_data.col_width_error_list|length }} 欄</span>
                            {% endif %}
                        </div>

                        <!-- 結構探測器 (含欄寬完整對照表) -->
                        <div class="inspector-box">
                            <button type="button" class="inspector-toggle" onclick="toggleInspector('{{ s_idx }}')">
                                🔍 展開/收合「{{ sheet_name }}」各欄欄寬對照、座標與公式結構表
                            </button>
                            <div class="inspector-content" id="inspector-{{ s_idx }}">
                                <p style="margin:0 0 6px 0; font-weight:bold; color:#2c3e50;">📏 本表欄寬度檢測對照表 (偏移補償：{{ offset_val }}、容許誤差：±{{ tol_val }})</p>
                                <table class="inspector-table">
                                    <tr>
                                        <th>欄位</th><th>Excel 換算欄寬</th><th>XML 原始值</th><th>預設標準欄寬</th><th>檢核狀態</th>
                                    </tr>
                                    {% for cw in s_data.col_width_inspector %}
                                    <tr>
                                        <td><strong>{{ cw.col }} 欄</strong></td>
                                        <td>{{ cw.est_display }}</td>
                                        <td style="color:#7f8c8d;">{{ cw.raw_xml }}</td>
                                        <td>{{ cw.preset }}</td>
                                        <td>
                                            {% if cw.is_error %}
                                                <span style="color:#c0392b; font-weight:bold;">❌ 不符</span>
                                            {% else %}
                                                <span style="color:#27ae60; font-weight:bold;">✅ 正常</span>
                                            {% endif %}
                                        </td>
                                    </tr>
                                    {% endfor %}
                                </table>

                                <p style="margin:8px 0 6px 0; font-weight:bold; color:#2c3e50;">📋 儲存格座標與原始公式 (共 {{ s_data.structure_info|length }} 格有值)</p>
                                <table class="inspector-table">
                                    <tr>
                                        <th>座標</th><th>顯示數值</th><th>原始公式 / 內容</th><th>字體</th><th>合併範圍</th>
                                    </tr>
                                    {% for item in s_data.structure_info %}
                                    <tr>
                                        <td><strong>{{ item.coord }}</strong></td>
                                        <td>{{ item.val }}</td>
                                        <td style="color:#2980b9; font-family:monospace;">{{ item.formula }}</td>
                                        <td>{{ item.font_size }}pt</td>
                                        <td>{{ item.merge_info }}</td>
                                    </tr>
                                    {% endfor %}
                                </table>
                            </div>
                        </div>

                        {% if s_data.excess_rows %}
                            <p style="color: #c0392b; font-weight: bold; background: #fadbd8; padding: 10px; border-radius: 5px; font-size: 11pt;">⚠️ 注意事項：本工作表未設定列印範圍且資料超過 {{ s_data.max_row }} 列，系統僅校對前 {{ s_data.max_row }} 列。</p>
                        {% endif %}

                        <!-- 欄寬異常專屬警示表 -->
                        {% if s_data.col_width_error_list %}
                        <div class="width-panel">
                            <h4>📏 發現欄位寬度與預設標準不符：</h4>
                            <table class="width-table">
                                <tr><th>欄位</th><th>目前欄寬 (Excel 換算)</th><th>XML 原始值</th><th>預設標準欄寬</th><th>差異值</th></tr>
                                {% for cw in s_data.col_width_error_list %}
                                <tr>
                                    <td><strong>{{ cw.col }} 欄</strong></td>
                                    <td style="color:#c0392b; font-weight:bold;">{{ cw.actual }}</td>
                                    <td style="color:#7f8c8d;">{{ cw.raw_xml }}</td>
                                    <td style="color:#27ae60; font-weight:bold;">{{ cw.preset }}</td>
                                    <td style="color:#c0392b;">{{ cw.diff }}</td>
                                </tr>
                                {% endfor %}
                            </table>
                        </div>
                        {% endif %}

                        {% if s_data.correction_list %}
                        <div class="correction-panel">
                            <h4>⚠️ 發現「跨表核對 / 小計合計 / 評點加總 / 面積算式」與系統運算不符：</h4>
                            <table class="correction-table">
                                <tr><th>位置</th><th>檢核項目</th><th>上傳檔案數值</th><th>系統校驗正確數值</th></tr>
                                {% for c in s_data.correction_list %}
                                <tr>
                                    <td>第 {{ c.row }} 列 ({{ c.coord }})</td>
                                    <td>{{ c.type }}</td>
                                    <td style="color:#c0392b; text-decoration: line-through;">{{ c.original }}</td>
                                    <td style="color:#27ae60; font-weight:bold; font-size:15px;">{{ c.correct }}</td>
                                </tr>
                                {% endfor %}
                            </table>
                        </div>
                        {% endif %}
                        
                        {% if not has_err %}
                            <p class="success">✅ 完美！本工作表所有檢核項目皆完全符合規定。</p>
                        {% else %}
                            <p style="margin-bottom:8px;">💡 滑鼠移至<span style="color:#c0392b; font-weight:bold;">紅色欄名／紅色或紫色格子</span>可查看錯誤原因（共發現 {{ s_data.error_list|length }} 處異常）。</p>
                        {% endif %}

                        <!-- 視覺化圖面 (主表顯示 B~L, N~Q；評點表顯示 D~M) -->
                        <div class="map-container">
                            <table class="excel-map">
                                <tr>
                                    <th class="row-num"></th>
                                    {% for c in s_data.visible_cols %}
                                        {% set cw = s_data.col_width_status.get(c, {'is_error': False, 'error': ''}) %}
                                        <th class="{% if cw.is_error %}col-width-error{% endif %}">
                                            {{ get_col_letter(c) }}{% if cw.is_error %} ⚠️{% endif %}
                                            {% if cw.is_error %}
                                                <div class="tooltip">{{ cw.error }}</div>
                                            {% endif %}
                                        </th>
                                    {% endfor %}
                                </tr>
                                {% for r in range(1, s_data.max_row + 1) %}
                                <tr>
                                    <th class="row-num">{{ r }}</th>
                                    {% for c in s_data.visible_cols %}
                                        {% set cell = s_data.map_data.get((r, c), {'status': 'empty', 'text': '', 'error': ''}) %}
                                        <td class="{% if cell.status == 'special' %}special-warning{% else %}cell-{{ cell.status }}{% endif %}">
                                            {{ cell.text }}
                                            {% if cell.status == 'error' or cell.status == 'special' %}
                                                <div class="tooltip">{{ cell.error }}</div>
                                            {% endif %}
                                        </td>
                                    {% endfor %}
                                </tr>
                                {% endfor %}
                            </table>
                        </div>

                    </div>
                </div>
                {% endfor %}
            </div>
            {% endif %}
        </div>

        <div class="history-panel">
            <div class="panel-actions">
                <button type="button" class="action-btn btn-top" onclick="scrollToTop()">⬆️ 置頂</button>
                <button type="button" class="action-btn btn-clear" onclick="confirmClearHistory()">🗑️ 清除紀錄</button>
            </div>
            <div class="history-header" onclick="toggleHistory()">
                <span>📂 瀏覽紀錄</span><span>▼</span>
            </div>
            <div class="history-content" id="historyContent">
                {% for uid, hist in history.items()|reverse %}
                    <div class="history-item {% if uid == current_id %}active{% endif %}">
                        <div onclick="window.location.href='/?id={{ uid }}'" style="cursor: pointer; font-weight:bold; color: #2980b9;">
                            {{ hist.filename }} <br>
                            <span class="history-time">{{ hist.time }}</span>
                        </div>
                        <ul style="margin: 6px 0 0 0; padding-left: 16px; font-size: 12px; color: #555;">
                            {% for s_name, s_data in hist.sheets.items() %}
                                <li onclick="window.location.href='/?id={{ uid }}#sheet-{{ loop.index }}'" style="cursor: pointer; padding: 2px 0; text-decoration: underline;">
                                    {% if s_data.sheet_type == 'rating' %}[評]{% endif %}{{ s_name }} 
                                    {% if s_data.error_list|length == 0 and not s_data.correction_list and not s_data.col_width_error_list %}
                                        <span style="color: #27ae60; font-weight: bold;">(✅)</span>
                                    {% else %}
                                        <span style="color: #c0392b; font-weight: bold;">(❌ {{ s_data.error_list|length }})</span>
                                    {% endif %}
                                </li>
                            {% endfor %}
                        </ul>
                    </div>
                {% else %}
                    <div style="padding: 14px; text-align: center; color: #999; font-size: 12px;">尚無紀錄</div>
                {% endfor %}
            </div>
        </div>
    </div>
</body>
</html>
"""

# ==============================================================================
# 🏢 處理「建物評點表」獨立檢核 (第 7~33 列納入加總)
# ==============================================================================
def process_rating_sheet(ws_data, ws_formula):
    visible_cols = BUILDING_RULES["rating_visible_cols"]  # D~M (4~13)
    error_list = []
    map_data = {}
    correction_list = []
    structure_info = []

    col_width_status, col_width_error_list, col_width_inspector = check_column_widths(
        ws_formula, visible_cols, RATING_PRESET_COL_WIDTHS
    )
    for cw_err in col_width_error_list:
        error_list.append(f"❌ {cw_err['msg']}")

    print_area = ws_data.print_area
    print_area_str = str(print_area) if print_area else "未設定"
    parsed_max_row = extract_max_row_from_print_area(print_area)

    max_row = max(parsed_max_row or 38, 38)
    merged_ranges = list(ws_data.merged_cells.ranges)

    all_calc_errors = {r: {c: [] for c in range(1, 21)} for r in range(1, max_row + 1)}
    active_blocks = 0

    # 依序檢核 4 棟建物區塊：F/G (6,7), H/I (8,9), J/K (10,11), L/M (12,13)
    for block_col in [6, 8, 10, 12]:
        score_col = block_col + 1
        col_letter = openpyxl.utils.get_column_letter(block_col)
        score_col_letter = openpyxl.utils.get_column_letter(score_col)

        r5_val = get_cell_value_safe(ws_data, 5, block_col, merged_ranges)
        has_r5 = (r5_val is not None and str(r5_val).strip() != "")

        has_any_score, expected_34 = sum_rating_block_scores(ws_data, score_col)
        r34_raw = get_cell_value_safe(ws_data, 34, block_col, merged_ranges)
        has_nonzero_34 = is_pure_number(r34_raw) and abs(to_float(r34_raw)) > 0.001

        r37_formula = get_cell_value_safe(ws_formula, 37, block_col, merged_ranges)
        r37_data = get_cell_value_safe(ws_data, 37, block_col, merged_ranges)
        r37_target = r37_formula if r37_formula is not None else r37_data
        has_r37 = (r37_target is not None and str(r37_target).strip() != "")

        if has_r5 or has_any_score or has_nonzero_34 or has_r37:
            active_blocks += 1

        # 單位面積評點數 (F34/H34/J34/L34) = 下一欄 7~33 列加總至小數點第一位
        if has_r5 or has_any_score or has_nonzero_34:
            if not is_pure_number(r34_raw):
                err_msg = f"單位面積評點數空白或非數字，{score_col_letter}7:{score_col_letter}33 加總應為 {expected_34:.1f}"
                add_cell_error(all_calc_errors, 34, block_col, err_msg, merged_ranges)
                correction_list.append({
                    "row": 34,
                    "coord": f"{col_letter}34",
                    "type": f"評點加總 ({score_col_letter}7:{score_col_letter}33)",
                    "original": "(空白/非數字)",
                    "correct": f"{expected_34:.1f}"
                })
            else:
                actual_34 = to_float(r34_raw)
                if abs(actual_34 - expected_34) > 0.05:
                    err_msg = f"單位面積評點數錯誤，{score_col_letter}7:{score_col_letter}33 加總應為 {expected_34:.1f}"
                    add_cell_error(all_calc_errors, 34, block_col, err_msg, merged_ranges)
                    correction_list.append({
                        "row": 34,
                        "coord": f"{col_letter}34",
                        "type": f"評點加總 ({score_col_letter}7:{score_col_letter}33)",
                        "original": f"{actual_34:.1f}",
                        "correct": f"{expected_34:.1f}"
                    })

        # 面積 (F38/H38/J38/L38) = 第 37 列計算式總和至小數點第二位
        if has_r37:
            is_valid, eval_area, err_msg = validate_and_eval_formula(r37_target)
            if not is_valid:
                add_cell_error(all_calc_errors, 37, block_col, f"建物尺寸計算式不合理：{err_msg}", merged_ranges)
            elif eval_area is not None:
                expected_38 = custom_round(eval_area, 2)
                r38_raw = get_cell_value_safe(ws_data, 38, block_col, merged_ranges)
                r38_err_col = block_col
                if r38_raw is None:
                    alt_38 = ws_data.cell(row=38, column=score_col).value
                    if alt_38 is not None:
                        r38_raw = alt_38
                        r38_err_col = score_col
                r38_coord = f"{openpyxl.utils.get_column_letter(r38_err_col)}38"

                if not is_pure_number(r38_raw):
                    msg = f"面積空白或非數字，依 {col_letter}37 計算應為 {expected_38:.2f}"
                    add_cell_error(all_calc_errors, 38, r38_err_col, msg, merged_ranges)
                    correction_list.append({
                        "row": 38,
                        "coord": r38_coord,
                        "type": f"面積算式驗算 ({col_letter}37)",
                        "original": "(空白/非數字)",
                        "correct": f"{expected_38:.2f}"
                    })
                else:
                    actual_38 = to_float(r38_raw)
                    if abs(actual_38 - expected_38) > 0.005:
                        msg = f"面積驗算錯誤，依 {col_letter}37 計算應為 {expected_38:.2f}"
                        add_cell_error(all_calc_errors, 38, r38_err_col, msg, merged_ranges)
                        correction_list.append({
                            "row": 38,
                            "coord": r38_coord,
                            "type": f"面積算式驗算 ({col_letter}37)",
                            "original": f"{actual_38:.2f}",
                            "correct": f"{expected_38:.2f}"
                        })

    # 組合 D~M 欄視覺化網格，並檢核 D5:M35 字體大小皆為 12pt
    for row in range(1, max_row + 1):
        for col in visible_cols:
            cell_formula = ws_formula.cell(row=row, column=col)
            cell_data = ws_data.cell(row=row, column=col)
            coord = cell_formula.coordinate

            is_phantom = False
            merge_range_str = "無"
            for mr in merged_ranges:
                min_col, min_row, max_mr_col, mr_max_row = mr.bounds
                if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
                    if row == min_row and col == min_col:
                        merge_range_str = str(mr)
                    else:
                        is_phantom = True
                    break

            errors = all_calc_errors[row][col].copy()

            if is_phantom:
                if errors:
                    err_msg = "、".join(errors)
                    map_data[(row, col)] = {'status': 'error', 'text': '(合併)', 'error': err_msg}
                    error_list.append(f"❌ {coord}: {err_msg}")
                else:
                    map_data[(row, col)] = {'status': 'empty', 'text': ''}
                continue

            raw_display_src = cell_data.value if cell_data.value is not None else cell_formula.value
            is_blank_cell = (raw_display_src is None) or (str(raw_display_src).strip() == "")

            if is_blank_cell:
                if errors:
                    err_msg = "、".join(errors)
                    map_data[(row, col)] = {'status': 'error', 'text': '(空)', 'error': err_msg}
                    error_list.append(f"❌ {coord}: {err_msg}")
                else:
                    map_data[(row, col)] = {'status': 'empty', 'text': ''}
                continue

            font_size = cell_formula.font.size if cell_formula.font else None

            if 5 <= row <= 35 and 4 <= col <= 13:
                if font_size is None or abs(float(font_size) - 12.0) > 0.01:
                    errors.append(f"D5:M35 字體應為 12pt (目前 {font_size}pt)")

            val_display = format_cell_display(raw_display_src, col, row, sheet_type="rating")
            formula_str = str(cell_formula.value if cell_formula.value is not None else "").strip()

            if len(structure_info) < 250:
                structure_info.append({
                    "coord": coord,
                    "val": str(cell_data.value if cell_data.value is not None else "")[:25],
                    "formula": formula_str[:35],
                    "font_size": font_size,
                    "merge_info": merge_range_str,
                })

            if errors:
                err_msg = "、".join(errors)
                map_data[(row, col)] = {'status': 'error', 'text': val_display, 'error': err_msg}
                error_list.append(f"❌ {coord}: {err_msg}")
            else:
                map_data[(row, col)] = {'status': 'valid', 'text': val_display}

    return {
        'sheet_type': 'rating',
        'visible_cols': visible_cols,
        'map_data': map_data,
        'error_list': error_list,
        'correction_list': correction_list,
        'col_width_status': col_width_status,
        'col_width_error_list': col_width_error_list,
        'col_width_inspector': col_width_inspector,
        'structure_info': structure_info,
        'print_area_str': print_area_str,
        'max_row': max_row,
        'excess_rows': False,
        'meta': {
            'active_blocks': active_blocks,
        }
    }

# ==============================================================================
# 📊 背景處理工作緒 (主估價表＋評點表＋跨表核對)
# ==============================================================================
def background_process_excel(task_id, filepath, original_filename):
    try:
        gc.enable()

        wb_data = openpyxl.load_workbook(filepath, data_only=True)
        wb_formula = openpyxl.load_workbook(filepath, data_only=False)

        sheets_data = {}
        skipped_sheets = []
        all_sheet_names = wb_data.sheetnames

        active_sheet_names = []
        rating_sheet_names = []
        for s_name in all_sheet_names:
            if should_skip_sheet(s_name):
                skipped_sheets.append(s_name)
                continue
            active_sheet_names.append(s_name)
            if is_rating_sheet(s_name):
                rating_sheet_names.append(s_name)

        total_sheets = max(len(active_sheet_names), 1)
        default_max_row = BUILDING_RULES["default_max_row"]
        data_start_row = BUILDING_RULES["data_start_row"]
        main_visible_cols = BUILDING_RULES["visible_cols"]

        # 預先建立去換行/空白版本的 BUILDING_DB 索引
        normalized_db = {}
        for (c_k, d_k), info in BUILDING_DB.items():
            norm_key = (
                str(c_k).replace('\n', '').replace('\r', '').strip(),
                str(d_k).replace('\n', '').replace('\r', '').strip()
            )
            normalized_db[norm_key] = info

        for idx, sheet_name in enumerate(active_sheet_names):
            progress = int((idx / total_sheets) * 100)
            tasks_db[task_id]['progress'] = progress
            tasks_db[task_id]['current_sheet'] = sheet_name

            ws_data = wb_data[sheet_name]
            ws_formula = wb_formula[sheet_name]

            # 若為「建物評點表」，執行獨立評點表檢核邏輯
            if is_rating_sheet(sheet_name):
                sheets_data[sheet_name] = process_rating_sheet(ws_data, ws_formula)
                gc.collect()
                continue

            # ================= 以下為「建物估價表 (主表)」檢核 =================
            error_list = []
            map_data = {}
            correction_list = []
            structure_info = []

            col_width_status, col_width_error_list, col_width_inspector = check_column_widths(
                ws_formula, main_visible_cols, PRESET_COL_WIDTHS
            )
            for cw_err in col_width_error_list:
                error_list.append(f"❌ {cw_err['msg']}")

            excess_rows = False
            print_area = ws_data.print_area
            print_area_str = str(print_area) if print_area else "未設定"

            parsed_max_row = extract_max_row_from_print_area(print_area)
            if parsed_max_row:
                max_row = parsed_max_row
            else:
                max_row = min(ws_data.max_row, default_max_row)
                excess_rows = ws_data.max_row > default_max_row

            merged_ranges = list(ws_data.merged_cells.ranges)

            all_calc_errors = {r: {c: [] for c in range(1, 21)} for r in range(1, max_row + 1)}
            all_special_flags = {r: {c: False for c in range(1, 21)} for r in range(1, max_row + 1)}

            c3_owner_name = str(get_cell_value_safe(ws_data, 3, 3, merged_ranges) or "").strip()
            matched_rating_name = find_matching_rating_sheet(sheet_name, c3_owner_name, rating_sheet_names)

            h1_raw = get_cell_value_safe(ws_data, 1, 8, merged_ranges)
            h1_str = str(h1_raw or "").strip()
            is_public_land = ("公有地" in h1_str)

            s24_raw = get_cell_value_safe(ws_data, 24, 19, merged_ranges)
            s24_str = str(s24_raw if s24_raw is not None else "").strip()
            s24_upper = s24_str.upper()

            s24_is_leased_or_false = (
                (s24_raw is False)
                or (s24_upper == "FALSE")
                or ("承租" in s24_str)
            )

            # 驗算 S 欄「合計」之 T 欄加總 (T2:合計上一列) 並比對 Q5
            q5_raw = get_cell_value_safe(ws_data, 5, 17, merged_ranges)
            q5_num = to_float(q5_raw) if is_pure_number(q5_raw) else None
            st_total_text = "未偵測"

            scan_st_max = min(max(max_row, ws_data.max_row), 200)
            for r_s in range(2, scan_st_max + 1):
                s_val = str(ws_data.cell(row=r_s, column=19).value or "").strip()
                if "合計" in s_val:
                    calc_t_sum = 0.0
                    for r_t in range(2, r_s):
                        t_item = ws_data.cell(row=r_t, column=20).value
                        if is_pure_number(t_item):
                            calc_t_sum += to_float(t_item)
                    calc_t_sum = custom_round(calc_t_sum, 2)

                    t_total_raw = ws_data.cell(row=r_s, column=20).value
                    t_total_num = to_float(t_total_raw) if is_pure_number(t_total_raw) else None
                    st_total_text = f"T{r_s}={t_total_num:,.2f} (重算={calc_t_sum:,.2f})" if t_total_num is not None else f"T{r_s}=空白 (重算={calc_t_sum:,.2f})"

                    if t_total_num is None or abs(t_total_num - calc_t_sum) > 0.01:
                        orig_disp = f"{t_total_num:,.2f}" if t_total_num is not None else "(空白/非數字)"
                        correction_list.append({
                            "row": r_s,
                            "coord": f"T{r_s}",
                            "type": "S欄合計 (T2至上一列加總)",
                            "original": orig_disp,
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                        error_list.append(f"❌ T{r_s}: S欄合計對應之T欄數值({orig_disp})與 T2:T{r_s-1} 加總({calc_t_sum:,.2f})不符")

                    if q5_num is None:
                        if 5 <= max_row:
                            add_cell_error(all_calc_errors, 5, 17, f"Q5面積為空白或非數字，應與 T{r_s} 加總 ({calc_t_sum:,.2f}) 相符", merged_ranges)
                        correction_list.append({
                            "row": 5,
                            "coord": "Q5",
                            "type": "Q5面積與T欄合計比對",
                            "original": "(空白/非數字)",
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                    elif abs(q5_num - calc_t_sum) > 0.01 or (t_total_num is not None and abs(q5_num - t_total_num) > 0.01):
                        if 5 <= max_row:
                            add_cell_error(
                                all_calc_errors, 5, 17,
                                f"Q5數值({q5_num:,.2f})與T欄合計(T{r_s}={t_total_num if t_total_num is not None else '空'} / 重算={calc_t_sum:,.2f})不符",
                                merged_ranges
                            )
                        correction_list.append({
                            "row": 5,
                            "coord": "Q5",
                            "type": "Q5面積與T欄合計比對",
                            "original": f"{q5_num:,.2f}",
                            "correct": f"{calc_t_sum:,.2f}"
                        })
                    break

            current_block_n_sum = 0.0
            subtotal_list = []
            has_subtotal_seen = False
            rating_block_idx = 0  # 第 1 棟=0(F欄), 第 2 棟=1(H欄), 第 3 棟=2(J欄), 第 4 棟=3(L欄)
            used_rating_sheet_label = None

            # 逐列檢核 (第 11 列 ~ max_row)
            for row in range(1, max_row + 1):
                if row < data_start_row:
                    continue

                col_B_val = str(ws_data.cell(row=row, column=2).value or "").strip()
                col_C_val = str(ws_data.cell(row=row, column=3).value or "").strip()
                col_D_val = str(ws_data.cell(row=row, column=4).value or "").strip()
                col_O_val = str(get_cell_value_safe(ws_data, row, 15, merged_ranges) or "").strip()

                row_text_sample = "".join(
                    str(ws_data.cell(row=row, column=c).value or "") for c in [2, 3, 12, 14]
                )
                is_footer_row = any(kw in row_text_sample for kw in ["注意事項", "調查單位", "調查日期", "以下空白"])

                is_subtotal_row = ("小計" in col_B_val)
                is_total_row = ("合計" in col_B_val) and not is_subtotal_row

                # ---------------- 【小計與合計驗算】 ----------------
                if is_subtotal_row:
                    n_raw = ws_data.cell(row=row, column=14).value
                    n_val_num = to_float(n_raw) if is_pure_number(n_raw) else 0.0
                    expected_sub = custom_round(current_block_n_sum, 0)

                    if abs(n_val_num - expected_sub) > 0.5:
                        correction_list.append({
                            "row": row,
                            "coord": f"N{row}",
                            "type": f"B{row} {col_B_val}",
                            "original": f"{n_val_num:,.0f}",
                            "correct": f"{expected_sub:,.0f}"
                        })
                        add_cell_error(all_calc_errors, row, 14, f"小計金額錯誤，上方區段加總應為 {expected_sub:,}", merged_ranges)

                    subtotal_list.append(expected_sub)
                    has_subtotal_seen = True
                    current_block_n_sum = 0.0
                    continue

                if is_total_row:
                    n_raw = ws_data.cell(row=row, column=14).value
                    n_val_num = to_float(n_raw) if is_pure_number(n_raw) else 0.0

                    if has_subtotal_seen:
                        expected_tot = custom_round(sum(subtotal_list), 0)
                        tot_desc = "各小計加總"
                    else:
                        expected_tot = custom_round(current_block_n_sum, 0)
                        tot_desc = f"N11:N{row - 1} 加總"

                    if abs(n_val_num - expected_tot) > 0.5:
                        correction_list.append({
                            "row": row,
                            "coord": f"N{row}",
                            "type": f"B{row} {col_B_val} ({tot_desc})",
                            "original": f"{n_val_num:,.0f}",
                            "correct": f"{expected_tot:,.0f}"
                        })
                        add_cell_error(all_calc_errors, row, 14, f"合計金額錯誤，{tot_desc}應為 {expected_tot:,}", merged_ranges)

                    current_block_n_sum = 0.0
                    subtotal_list = []
                    has_subtotal_seen = False
                    continue

                if is_footer_row:
                    continue

                n_cell_val = ws_data.cell(row=row, column=14).value
                if is_pure_number(n_cell_val):
                    current_block_n_sum += to_float(n_cell_val)

                has_col_B_id = (
                    col_B_val != ""
                    and not any(kw in col_B_val for kw in ["編號", "項目", "小計", "合計", "總計", "注意", "調查", "以下空白", "備註"])
                )

                # ---------------- 【當 F 欄存在數字時：跨表比對評點表 (第 7~33 列加總)】 ----------------
                f_raw = ws_data.cell(row=row, column=6).value
                has_f_num = is_pure_number(f_raw)

                if has_f_num:
                    if not matched_rating_name:
                        add_cell_error(
                            all_calc_errors, row, 6,
                            f"F{row}有評點單價，但找不到名稱包含「{sheet_name}」的建物評點表",
                            merged_ranges
                        )
                    elif rating_block_idx >= 4:
                        add_cell_error(
                            all_calc_errors, row, 6,
                            "評點建物數量超過評點表 4 棟欄位上限 (F/H/J/L)",
                            merged_ranges
                        )
                    else:
                        used_rating_sheet_label = matched_rating_name
                        ws_r_data = wb_data[matched_rating_name]
                        ws_r_formula = wb_formula[matched_rating_name]
                        r_merges = list(ws_r_data.merged_cells.ranges)

                        target_col = 6 + (rating_block_idx * 2)  # F=6, H=8, J=10, L=12
                        score_col = target_col + 1               # G=7, I=9, K=11, M=13
                        col_letter = openpyxl.utils.get_column_letter(target_col)
                        score_col_letter = openpyxl.utils.get_column_letter(score_col)

                        # (A) 比對主表 B 欄編號 vs 評點表第 5 列編號 (F5/H5/J5/L5)
                        main_id = normalize_id_str(ws_data.cell(row=row, column=2).value)
                        r5_id = normalize_id_str(get_cell_value_safe(ws_r_data, 5, target_col, r_merges))
                        if main_id != r5_id:
                            add_cell_error(
                                all_calc_errors, row, 2,
                                f"編號({main_id or '空'})與「{matched_rating_name}」{col_letter}5 ({r5_id or '空'}) 不符",
                                merged_ranges
                            )
                            correction_list.append({
                                "row": row,
                                "coord": f"B{row}",
                                "type": f"跨表核對編號 ({matched_rating_name} {col_letter}5)",
                                "original": main_id or "(空)",
                                "correct": r5_id or "(空)"
                            })

                        # (B) 比對主表 E 欄 vs 評點表 F34/H34/J34/L34 (並同步檢查 7~33 列重算總和)
                        has_scores, calc_pts = sum_rating_block_scores(ws_r_data, score_col)
                        r34_raw = get_cell_value_safe(ws_r_data, 34, target_col, r_merges)
                        r34_num = custom_round(to_float(r34_raw), 1) if is_pure_number(r34_raw) else None
                        expected_e = calc_pts if has_scores else r34_num

                        e_raw = ws_data.cell(row=row, column=5).value
                        if expected_e is None:
                            add_cell_error(
                                all_calc_errors, row, 5,
                                f"「{matched_rating_name}」{col_letter}34 及 {score_col_letter}7:{score_col_letter}33 無評點數值",
                                merged_ranges
                            )
                        else:
                            e_is_mismatch = (not is_pure_number(e_raw)) or (abs(to_float(e_raw) - expected_e) > 0.05)
                            r34_differs_from_e = (r34_num is not None) and is_pure_number(e_raw) and (abs(to_float(e_raw) - r34_num) > 0.05)
                            if e_is_mismatch or r34_differs_from_e:
                                detail_note = f"{col_letter}34={r34_num if r34_num is not None else '空'} / 重算={expected_e:.1f}"
                                add_cell_error(
                                    all_calc_errors, row, 5,
                                    f"與「{matched_rating_name}」評點數不符 ({detail_note})",
                                    merged_ranges
                                )
                                correction_list.append({
                                    "row": row,
                                    "coord": f"E{row}",
                                    "type": f"跨表核對評點 ({matched_rating_name} {col_letter}34)",
                                    "original": str(e_raw) if e_raw is not None else "(空)",
                                    "correct": f"{expected_e:.1f}"
                                })

                        # (C) 比對主表 G 欄 vs 評點表 F37/H37/J37/L37 計算式結果 (小數點第 2 位)
                        r37_formula = get_cell_value_safe(ws_r_formula, 37, target_col, r_merges)
                        r37_data = get_cell_value_safe(ws_r_data, 37, target_col, r_merges)
                        r37_target = r37_formula if r37_formula is not None else r37_data
                        g_raw = ws_data.cell(row=row, column=7).value

                        if r37_target is None or str(r37_target).strip() == "":
                            add_cell_error(
                                all_calc_errors, row, 7,
                                f"「{matched_rating_name}」{col_letter}37 建物尺寸為空白，無法核對面積",
                                merged_ranges
                            )
                        else:
                            r37_valid, r37_area, r37_err = validate_and_eval_formula(r37_target)
                            if not r37_valid:
                                add_cell_error(
                                    all_calc_errors, row, 7,
                                    f"「{matched_rating_name}」{col_letter}37 計算式異常：{r37_err}",
                                    merged_ranges
                                )
                            elif r37_area is not None:
                                expected_g = custom_round(r37_area, 2)
                                if not is_pure_number(g_raw) or abs(to_float(g_raw) - expected_g) > 0.005:
                                    add_cell_error(
                                        all_calc_errors, row, 7,
                                        f"與「{matched_rating_name}」{col_letter}37 計算面積 ({expected_g:.2f}) 不符",
                                        merged_ranges
                                    )
                                    correction_list.append({
                                        "row": row,
                                        "coord": f"G{row}",
                                        "type": f"跨表核對面積 ({matched_rating_name} {col_letter}37)",
                                        "original": str(g_raw) if g_raw is not None else "(空)",
                                        "correct": f"{expected_g:.2f}"
                                    })
                    rating_block_idx += 1

                # ---------------- 【C 欄 & D 欄 資料庫檢核 H 欄單價、K 欄單位】 ----------------
                norm_key = (
                    col_C_val.replace('\n', '').replace('\r', '').strip(),
                    col_D_val.replace('\n', '').replace('\r', '').strip()
                )
                if norm_key in normalized_db:
                    db_item = normalized_db[norm_key]
                    h_raw = ws_data.cell(row=row, column=8).value
                    k_raw = str(ws_data.cell(row=row, column=11).value or "").strip()

                    if not is_pure_number(h_raw):
                        add_cell_error(all_calc_errors, row, 8, f"單價空白或非數字，資料庫標準應為 {db_item['price']}", merged_ranges)
                    elif abs(to_float(h_raw) - db_item["price"]) > 0.001:
                        add_cell_error(all_calc_errors, row, 8, f"單價與資料庫不符，應為 {db_item['price']}", merged_ranges)

                    if k_raw != db_item["unit"]:
                        add_cell_error(all_calc_errors, row, 11, f"單位與資料庫不符，應為「{db_item['unit']}」", merged_ranges)

                # ---------------- 【I 欄計算式檢核與 J 欄四捨五入驗算】 ----------------
                i_formula_raw = ws_formula.cell(row=row, column=9).value
                i_data_raw = ws_data.cell(row=row, column=9).value
                i_check_target = i_formula_raw if i_formula_raw is not None else i_data_raw

                if i_check_target is not None and str(i_check_target).strip() != "":
                    i_valid, i_eval_val, i_err_msg = validate_and_eval_formula(i_check_target)
                    if not i_valid:
                        add_cell_error(all_calc_errors, row, 9, i_err_msg, merged_ranges)
                    elif i_eval_val is not None:
                        expected_j = custom_round(i_eval_val, ndigits=2)
                        j_raw = ws_data.cell(row=row, column=10).value
                        if not is_pure_number(j_raw):
                            add_cell_error(all_calc_errors, row, 10, f"數量空白或非數字，依 I 欄計算應為 {expected_j}", merged_ranges)
                        else:
                            j_num = to_float(j_raw)
                            if abs(j_num - expected_j) > 0.005:
                                add_cell_error(all_calc_errors, row, 10, f"數量驗算錯誤，I 欄四捨五入後應為 {expected_j}", merged_ranges)

                # ---------------- 【H1 公有地與 S24 連動檢核 L 欄】 ----------------
                if is_public_land and has_col_B_id:
                    has_illegal_or_no_water = ("無合法" in col_O_val) or ("無水權" in col_O_val)
                    if s24_is_leased_or_false:
                        if has_illegal_or_no_water:
                            expected_l = 0.7
                            l_reason = "O欄含無合法/無水權，L欄應為 0.7 (70%)"
                        else:
                            expected_l = 1.0
                            l_reason = "H1含公有地且S24含承租/FALSE，L欄應為 1.0 (100%)"
                    else:
                        expected_l = 0.5
                        l_reason = "H1含公有地且S24不含承租/TRUE，L欄應為 0.5 (50%)"

                    l_raw = ws_data.cell(row=row, column=12).value
                    l_rate_parsed = parse_rate_value(l_raw)
                    if l_rate_parsed is None:
                        add_cell_error(all_calc_errors, row, 12, f"發放比例空白或非數字 ({l_reason})", merged_ranges)
                    elif abs(l_rate_parsed - expected_l) > 0.001:
                        add_cell_error(all_calc_errors, row, 12, f"發放比例錯誤 ({l_reason})", merged_ranges)

                # ---------------- 【N 欄金額驗算】 ----------------
                l_raw = ws_data.cell(row=row, column=12).value
                if has_f_num:
                    # 需求 1：當 F 欄存在數字時，N = round(round(E * F * G, 0) * L, 0)
                    e_raw_n = ws_data.cell(row=row, column=5).value
                    g_raw_n = ws_data.cell(row=row, column=7).value
                    l_num = parse_rate_value(l_raw)

                    if not is_pure_number(e_raw_n) or not is_pure_number(g_raw_n):
                        add_cell_error(
                            all_calc_errors, row, 14,
                            "無法驗算房屋現值：E 欄(點/㎡)或 G 欄(評點房屋面積)空白或非數字",
                            merged_ranges
                        )
                    elif l_num is None:
                        add_cell_error(
                            all_calc_errors, row, 14,
                            "無法驗算房屋現值：L 欄發放比例空白或非數字",
                            merged_ranges
                        )
                    else:
                        e_num_n = to_float(e_raw_n)
                        f_num_n = to_float(f_raw)
                        g_num_n = to_float(g_raw_n)
                        step1 = custom_round(e_num_n * f_num_n * g_num_n, ndigits=0)
                        expected_n = custom_round(step1 * l_num, ndigits=0)

                        if not is_pure_number(n_cell_val):
                            add_cell_error(
                                all_calc_errors, row, 14,
                                f"金額空白或非數字，應為 {expected_n:,} (E×F×G四捨五入={step1:,}，再×L四捨五入)",
                                merged_ranges
                            )
                        else:
                            actual_n = to_float(n_cell_val)
                            if abs(actual_n - expected_n) > 0.1:
                                add_cell_error(
                                    all_calc_errors, row, 14,
                                    f"金額驗算錯誤，應為 {expected_n:,} (E×F×G四捨五入={step1:,}，再×L四捨五入)",
                                    merged_ranges
                                )
                else:
                    # 一般地上物列：N = round(round(H * J, 0) * L, 0)
                    h_raw = ws_data.cell(row=row, column=8).value
                    j_raw = ws_data.cell(row=row, column=10).value

                    if is_pure_number(h_raw) and is_pure_number(j_raw):
                        h_num = to_float(h_raw)
                        j_num = to_float(j_raw)
                        l_num = parse_rate_value(l_raw)
                        if l_num is None:
                            add_cell_error(all_calc_errors, row, 14, "無法驗算金額：L 欄發放比例空白或非數字", merged_ranges)
                        else:
                            step1 = custom_round(h_num * j_num, ndigits=0)
                            expected_n = custom_round(step1 * l_num, ndigits=0)
                            if not is_pure_number(n_cell_val):
                                add_cell_error(all_calc_errors, row, 14, f"金額空白或非數字，應為 {expected_n:,}", merged_ranges)
                            else:
                                actual_n = to_float(n_cell_val)
                                if abs(actual_n - expected_n) > 0.1:
                                    add_cell_error(
                                        all_calc_errors, row, 14,
                                        f"金額驗算錯誤，應為 {expected_n:,} (H×J四捨五入={step1:,}，再×L四捨五入)",
                                        merged_ranges
                                    )

            # 組合主表視覺化網格 (僅渲染 main_visible_cols: B~L, N~Q)
            for row in range(1, max_row + 1):
                for col in main_visible_cols:
                    cell_formula = ws_formula.cell(row=row, column=col)
                    cell_data = ws_data.cell(row=row, column=col)
                    coord = cell_formula.coordinate

                    is_phantom = False
                    merge_range_str = "無"
                    for mr in merged_ranges:
                        min_col, min_row, max_mr_col, mr_max_row = mr.bounds
                        if min_row <= row <= mr_max_row and min_col <= col <= max_mr_col:
                            if row == min_row and col == min_col:
                                merge_range_str = str(mr)
                            else:
                                is_phantom = True
                            break

                    errors = all_calc_errors[row][col].copy()
                    is_special = all_special_flags[row][col]

                    if is_phantom:
                        if errors:
                            err_msg = "、".join(errors)
                            map_data[(row, col)] = {'status': 'error', 'text': '(合併)', 'error': err_msg}
                            error_list.append(f"❌ {coord}: {err_msg}")
                        else:
                            map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    if cell_data.value is None and cell_formula.value is None:
                        if errors:
                            err_msg = "、".join(errors)
                            map_data[(row, col)] = {'status': 'error', 'text': '(空)', 'error': err_msg}
                            error_list.append(f"❌ {coord}: {err_msg}")
                        else:
                            map_data[(row, col)] = {'status': 'empty', 'text': ''}
                        continue

                    val_display = format_cell_display(
                        cell_data.value if cell_data.value is not None else cell_formula.value,
                        col,
                        row,
                        sheet_type="main"
                    )
                    formula_str = str(cell_formula.value if cell_formula.value is not None else "").strip()
                    font_size = cell_formula.font.size if cell_formula.font else None

                    if len(structure_info) < 250:
                        structure_info.append({
                            "coord": coord,
                            "val": str(cell_data.value if cell_data.value is not None else "")[:25],
                            "formula": formula_str[:35],
                            "font_size": font_size,
                            "merge_info": merge_range_str,
                        })

                    if errors:
                        err_msg = "、".join(errors)
                        status = 'special' if is_special else 'error'
                        map_data[(row, col)] = {'status': status, 'text': val_display, 'error': err_msg}
                        error_list.append(f"{'🟪' if is_special else '❌'} {coord}: {err_msg}")
                    else:
                        map_data[(row, col)] = {'status': 'valid', 'text': val_display}

            sheets_data[sheet_name] = {
                'sheet_type': 'main',
                'visible_cols': main_visible_cols,
                'map_data': map_data,
                'error_list': error_list,
                'correction_list': correction_list,
                'col_width_status': col_width_status,
                'col_width_error_list': col_width_error_list,
                'col_width_inspector': col_width_inspector,
                'structure_info': structure_info,
                'print_area_str': print_area_str,
                'max_row': max_row,
                'excess_rows': excess_rows,
                'meta': {
                    'h1_text': h1_str[:22],
                    's24_text': s24_str if s24_str else '(空)',
                    'q5_text': f"{q5_num:,.2f}" if q5_num is not None else str(q5_raw or '(空)'),
                    'st_total_text': st_total_text,
                    'matched_rating': used_rating_sheet_label or (f"{matched_rating_name} (本表無評點列)" if matched_rating_name else "無"),
                }
            }

            gc.collect()

        wb_data.close()
        wb_formula.close()
        del wb_data
        del wb_formula
        gc.collect()

        record_id = str(uuid.uuid4())[:8]
        history_db[record_id] = {
            'filename': original_filename,
            'time': time.strftime("%H:%M:%S"),
            'sheets': sheets_data,
            'skipped_sheets': skipped_sheets
        }

        tasks_db[task_id]['status'] = 'completed'
        tasks_db[task_id]['record_id'] = record_id

    except Exception as e:
        tasks_db[task_id]['status'] = 'error'
        tasks_db[task_id]['message'] = str(e)
    finally:
        if os.path.exists(filepath):
            try:
                os.remove(filepath)
            except Exception:
                pass


# --- 路由設定 ---

@app.route("/", methods=["GET"])
def index():
    record_id = request.args.get('id')
    common_ctx = dict(
        history=history_db,
        offset_val=BUILDING_RULES["col_width_xml_offset"],
        tol_val=BUILDING_RULES["col_width_tolerance"],
        get_col_letter=openpyxl.utils.get_column_letter
    )
    if record_id and record_id in history_db:
        return render_template_string(HTML_TEMPLATE, record=history_db[record_id], current_id=record_id, **common_ctx)
    return render_template_string(HTML_TEMPLATE, record=None, current_id=None, **common_ctx)

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files.get("excel_file")
    if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xlsm')):
        task_id = str(uuid.uuid4())

        temp_dir = tempfile.gettempdir()
        filepath = os.path.join(temp_dir, f"{task_id}_{file.filename}")
        file.save(filepath)

        tasks_db[task_id] = {'status': 'processing', 'progress': 0, 'current_sheet': ''}

        thread = threading.Thread(target=background_process_excel, args=(task_id, filepath, file.filename))
        thread.start()

        return jsonify({"task_id": task_id})
    return jsonify({"error": "不支援的檔案格式，請上傳 .xlsx 或 .xlsm 檔案"})

@app.route("/status/<task_id>", methods=["GET"])
def task_status(task_id):
    task = tasks_db.get(task_id)
    if task:
        return jsonify(task)
    return jsonify({"status": "error", "message": "找不到該任務"})

@app.route("/clear_history", methods=["POST"])
def clear_history():
    history_db.clear()
    finished_tasks = [tid for tid, t in tasks_db.items() if t.get('status') != 'processing']
    for tid in finished_tasks:
        del tasks_db[tid]
    gc.collect()
    return jsonify({"status": "cleared"})

if __name__ == "__main__":
    app.run(debug=True, port=5000)
import streamlit as st
import pandas as pd
from io import BytesIO
import datetime
import difflib
import json
import base64
import streamlit.components.v1 as components

# --- 1. 辅助工具函数 ---
def is_duration_col(col_name, custom_time_keywords=None):
    """智能识别时间/时长列"""
    include_k = custom_time_keywords if custom_time_keywords else ['总', '时长', '时间', '片长', '总长', '总片长', 'Duration', 'time']
    exclude_k = ['上线', '日期', '发布', '开播', '年度', '更新', 'date', 'day', '总集数', '总片数', '总集', 'ID', '号', 'No']
    is_inc = any(k in col_name for k in include_k)
    is_exc = any(k in col_name for k in exclude_k)
    return is_inc and not is_exc

def parse_time_logic(val):
    """支持超过 24 小时的时间数据强制解析"""
    if pd.isna(val) or val == "": return ""
    v_str = str(val).strip()
    if ':' in v_str:
        parts = v_str.split(':')
        try:
            if len(parts) == 3:
                return (int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])) / 86400.0
            elif len(parts) == 2:
                return (int(parts[0]) * 60 + int(parts[1])) / 84600.0
        except: pass
    try: return float(v_str)
    except: return v_str

def find_headers(df, custom_keywords=None):
    """智能侦测表头所在行"""
    best_idx, max_m = 0, -1
    core = custom_keywords if custom_keywords else {'编码', '名称', '日期', '类型', '状态', '标题', '时间', 'ID', 'Name', 'Date', 'Type'}
    
    for idx in range(min(10, len(df))):
        row = [str(x).strip() for x in df.iloc[idx].values if pd.notna(x)]
        m = sum(3 if any(c in v for c in core) else (0.8 if len(v)>1 else 0) for v in row)
        if len(row) > 0:
            num_count = sum(1 for v in row if v.replace('.','',1).isdigit())
            if num_count / len(row) > 0.5:
                m -= 5
        if m > max_m: max_m, best_idx = m, idx
        
    h = [str(x).strip() for x in df.iloc[best_idx].values]
    return best_idx, [x for x in h if x and x != 'nan' and not x.startswith('Unnamed:')]

def get_best_sheet_match(target_sheet, available_sheets):
    """
    强力模糊匹配 Sheet 名字：
    无视首尾空格、中英文括号差异、大小写差异
    """
    if target_sheet in available_sheets: return target_sheet
    
    def clean_str(s):
        return str(s).strip().lower().replace(" ", "").replace("（", "(").replace("）", ")").replace("-", "")
        
    target_clean = clean_str(target_sheet)
    
    # 1. 净化后精准匹配
    for s in available_sheets:
        if clean_str(s) == target_clean: return s
        
    # 2. 包含关系匹配 (比如 "4K少儿" 在 "4K少儿(最新)" 里)
    for s in available_sheets:
        if target_clean in clean_str(s) or clean_str(s) in target_clean: return s
        
    # 3. Difflib 相似度匹配
    matches = difflib.get_close_matches(target_clean, [clean_str(x) for x in available_sheets], n=1, cutoff=0.5)
    if matches:
        for s in available_sheets:
            if clean_str(s) == matches[0]: return s
            
    # 如果真的长得完全不一样，才兜底返回第一个
    return available_sheets[0]

# --- 2. 页面基本配置 ---
st.set_page_config(page_title="万能 Excel 表头提取与排版助手", layout="wide", page_icon="🔀")

# --- 3. 顶部说明区 ---
st.title("🔀 万能 Excel 表头提取与排版助手")

st.markdown("""
✅ 解决 Excel 表头/字段对应提取转换的问题，支持自动相似度对齐。

✅ 支持时间列强制累加，无视 24 小时进位限制。

✅ **强力多 Sheet 智能寻呼**：表A 与表B Sheet 名即使有细微差异（如空格、全半角括号），系统也会强力模糊对齐，拒绝数据串台！

---

系统会自动把 **【表A（原始数据）】** 里的数据提取/筛选并填入 **【表B（模版）】** 里，**最终导出的 Sheet 将完全沿用表B的结构与命名。**

* **自定义填充：** 如果A表没有B表某字段，你可以设置默认内容，不设置视为空。（左侧第 1 点）
* **无表头预警：** 发现无名列时系统会红框预警；可见左侧第2点手动修正。
* **求助与售后：** 遇到卡顿请点击右上角 `⋮` ➔ `Clear cache` ➔ `Rerun`。报错请通过下方一键复制按钮发给售后。
""")

# --- 4. 侧边栏：1. 个性化默认值填充 ---
st.sidebar.header("⚙️ 1. 个性化默认值填充")
st.sidebar.info("👉 **用法：** 表B字段名=你要填的内容。每行一个。未设置的字段为空。")
today_str = datetime.datetime.now().strftime("%y%m%d")
custom_defaults_text = st.sidebar.text_area("输入填充规则：", placeholder=f"更新日期={today_str}\n是否成品=是", height=150)

custom_defaults = {}
if custom_defaults_text.strip():
    for line in custom_defaults_text.split('\n'):
        if '=' in line:
            parts = line.split('=')
            if len(parts) == 2: custom_defaults[parts[0].strip()] = parts[1].strip()

# --- 5. 侧边栏：2. 手动对号修正 ---
st.sidebar.markdown("---")
st.sidebar.header("🛠️ 2. 手动对号修正")
st.sidebar.info("👉 **用法：** 表B字段名=表A列号。例如：UMAI=58。输入完按回车生效。")
manual_map_text = st.sidebar.text_area("输入映射规则（不区分大小写）：", placeholder="许可证=10", height=120)

manual_map_config = {}
if manual_map_text.strip():
    for line in manual_map_text.split('\n'):
        if '=' in line:
            parts = line.split('=')
            if len(parts) == 2:
                try: manual_map_config[parts[0].strip().lower()] = int(parts[1].strip()) - 1
                except ValueError: pass

# --- 6. 主界面：文件上传 ---
col_u1, col_u2 = st.columns(2)
with col_u1: raw_file = st.file_uploader("📂 上传【表A：数据源】(系统导出的原始表)", type=["csv", "xlsx", "xls"])
with col_u2: template_file = st.file_uploader("📋 上传【表B：目标模板】(带有目标格式的模版)", type=["xlsx", "xls"])

# ==================== 7. 读取表A结构 ====================
is_raw_excel = False
xls_raw = None
raw_sheet_names = []

if raw_file:
    try:
        ext = raw_file.name.split('.')[-1].lower()
        if ext in ['xlsx', 'xls']:
            is_raw_excel = True
            xls_raw = pd.ExcelFile(raw_file)
            raw_sheet_names = xls_raw.sheet_names
            st.success(f"✅ 表A 读取成功！包含 {len(raw_sheet_names)} 个 Sheet。")
        else:
            st.success("✅ 表A (CSV单表) 读取成功！")
    except Exception as e: 
        st.error(f"读取表A失败: {e}")
        st.stop()

# ==================== 8. 侧边栏：3. 自查与特征词自定义 ====================
selected_time_cols = []
force_header_config = {}

st.sidebar.markdown("---")
st.sidebar.header("🔍 3. 自查—表B表头行号修正")

header_keywords_input = st.sidebar.text_input(
    "🧠 自定义表头定位特征词 (逗号隔开)",
    value="编码,名称,日期,类型,状态,标题,时间,片单,导演,演员,ID,Name,Date,Type",
    help="系统会根据这些词在前10行中自动定位表头。"
)
custom_header_keywords = set(x.strip() for x in header_keywords_input.replace("，", ",").split(",") if x.strip())

time_keywords_input = st.sidebar.text_input(
    "⏳ 自定义时间列特征词 (逗号隔开)",
    value="总,时长,时间,片长,总长,总片长,Duration,time",
    help="表头只要包含这些词，就会强制 [h]:mm:ss 累加转换。"
)
custom_time_keywords = [x.strip() for x in time_keywords_input.replace("，", ",").split(",") if x.strip()]

if template_file:
    xls_tpl = pd.ExcelFile(template_file)
    for s_name in xls_tpl.sheet_names:
        force_header_config[s_name] = st.sidebar.number_input(
            f"『{s_name}』的字段名在表B第几行？", 
            min_value=0, max_value=20, value=0, 
            key=f"row_{s_name}"
        )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("⏳ 时间格式转换确认")
    all_time_cols_in_template = []
    for s_name in xls_tpl.sheet_names:
        try:
            df_preview = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=10)
            _, hs = find_headers(df_preview, custom_header_keywords)
            all_time_cols_in_template.extend([h for h in hs if is_duration_col(h, custom_time_keywords)])
        except: pass
    all_time_cols_in_template = list(set(all_time_cols_in_template))
    
    selected_time_cols = st.sidebar.multiselect(
        "以下字段将执行 [h]:mm:ss 累加转换", 
        all_time_cols_in_template, 
        default=all_time_cols_in_template
    )
else:
    st.sidebar.caption("⏳ 上传【表B】后，即可在此进行设置。")

# ==================== 9. 执行核心映射与渲染看板 ====================
if raw_file and template_file:
    output = BytesIO()
    final_reports = {}
    debug_log = {"诊断": {}}
    sheet_routing_log = {}

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for s_name in xls_tpl.sheet_names:
            df_tpl_meta = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=10).fillna("")
            if df_tpl_meta.empty: continue
            
            f_h = force_header_config.get(s_name, 0)
            h_idx, headers = (f_h-1, [str(x).strip() for x in df_tpl_meta.iloc[f_h-1].values if pd.notna(x)]) if f_h > 0 else find_headers(df_tpl_meta, custom_header_keywords)
            
            # --- 智能路由：寻找表 A 中对应的 Sheet ---
            target_raw_sheet = "CSV单表文件"
            if is_raw_excel:
                target_raw_sheet = get_best_sheet_match(s_name, raw_sheet_names)
                temp_raw = pd.read_excel(xls_raw, sheet_name=target_raw_sheet, header=None, nrows=15, dtype=str).fillna("")
                a_h_idx, _ = find_headers(temp_raw, custom_header_keywords)
                df_raw_current = pd.read_excel(xls_raw, sheet_name=target_raw_sheet, header=a_h_idx, dtype=str).fillna("")
            else:
                raw_file.seek(0)
                temp_raw = pd.read_csv(raw_file, header=None, nrows=15, dtype=str).fillna("")
                a_h_idx, _ = find_headers(temp_raw, custom_header_keywords)
                raw_file.seek(0)
                df_raw_current = pd.read_csv(raw_file, header=a_h_idx, dtype=str).fillna("")
            
            # 记录路由来源供UI展示
            sheet_routing_log[s_name] = target_raw_sheet

            # 无名列预警
            unnamed = []
            for i, col in enumerate(df_raw_current.columns, 1):
                col_str = str(col).strip()
                if "Unnamed" in col_str or col_str == "" or col_str.startswith("Column"):
                    sample = [str(x).strip() for x in df_raw_current.iloc[:, i-1].tolist() if str(x).strip() != ""]
                    unnamed.append((i, "、".join(sample[:3]) if sample else "全空"))
            
            if unnamed:
                st.warning(f"🚨 **表A ({target_raw_sheet}) 发现无名列！** 请根据预览在左侧指定列号：")
                for idx, pre in unnamed: st.write(f"👉 第 `{idx}` 列 ➔ 内容预览: `{pre}`")

            # 匹配与填充逻辑
            out_df = pd.DataFrame(columns=headers)
            raw_cols = df_raw_current.columns.tolist()
            report = []
            
            for b_idx, col_name in enumerate(headers, 1):
                a_idx, status = None, "empty"
                if col_name.lower() in manual_map_config:
                    a_idx = manual_map_config[col_name.lower()]; status = "ok"
                else:
                    m = difflib.get_close_matches(col_name, raw_cols, n=1, cutoff=0.4)
                    if m: a_idx = raw_cols.index(m[0]); status = "ok"
                
                if status == "ok" and a_idx < len(df_raw_current.columns):
                    series = df_raw_current.iloc[:, a_idx]
                    if col_name in selected_time_cols: series = series.apply(parse_time_logic)
                    out_df[col_name] = series
                    report.append({"b": b_idx, "bn": col_name, "a": a_idx + 1, "an": raw_cols[a_idx], "s": "ok"})
                else:
                    fill = custom_defaults.get(col_name, "")
                    out_df[col_name] = fill
                    report.append({"b": b_idx, "bn": col_name, "f": fill, "s": "fill" if fill else "empty"})
            
            # 写入并渲染时间格式
            out_df.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for i, h in enumerate(headers, 1):
                if h in selected_time_cols:
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 14
                    for r in range(2, len(out_df)+2):
                        cell = ws.cell(row=r, column=i)
                        if isinstance(cell.value, (float, int)): 
                            cell.number_format = '[h]:mm:ss'
            
            final_reports[s_name] = report
            debug_log["诊断"][s_name] = {"强制寻呼自": target_raw_sheet, "字段": headers}

    # ==================== 看板与下载区 ====================
    st.markdown("---")
    st.subheader("📊 数据对齐与填充看板")

    for s_name, report in final_reports.items():
        ok_count = sum(1 for x in report if x['s'] == 'ok')
        source_sheet = sheet_routing_log.get(s_name, "未知")
        
        # UI直白展示数据的真实来源
        with st.expander(f"📁 表B『{s_name}』 (数据溯源: 自动抓取表A『{source_sheet}』) | 提取成功: {ok_count} | 填充留空: {len(report)-ok_count}"):
            c1, c2 = st.columns(2)
            with c1:
                st.write("🟢 **表A 提取成功**")
                for x in [item for item in report if item['s'] == 'ok']:
                    st.write(f"A第`{x['a']}`列({x['an']}) ➔ B第`{x['b']}`列({x['bn']})")
                if not any(item['s'] == 'ok' for item in report):
                    st.caption("没有成功从表A提取的数据。")
            with c2:
                st.write("🟡 **表B 填充或空**")
                for x in [item for item in report if item['s'] != 'ok']:
                    val_text = f"填: `{x['f']}`" if x['f'] else "留空"
                    st.write(f"B第`{x['b']}`列({x['bn']}) ➔ {val_text}")

    st.success("🎉 匹配完成！所有数据已通过强化寻呼精准入座！")
    st.download_button("📥 下载转换后的交付 Excel", data=output.getvalue(), file_name=f"交付表_{today_str}.xlsx")
    
    # --- 10. 排错专区 ---
    st.markdown("---")
    with st.expander("🛠️ 排错专区 (遇到问题请复制此段)"):
        log_json = json.dumps(debug_log, ensure_ascii=False, indent=2)
        b64_json = base64.b64encode(log_json.encode('utf-8')).decode('utf-8')
        
        copy_html = f"""
        <div style="text-align: right; margin-bottom: -15px;">
            <button id="copy-btn" style="
                background-color: #ff4b4b; 
                color: white; 
                border: none; 
                padding: 6px 14px; 
                font-size: 13px; 
                border-radius: 4px; 
                cursor: pointer;
                font-family: inherit;
                font-weight: 500;
            ">📋 一键复制全部错误信息</button>
        </div>
        <script>
        function decodeB64(str) {{
            return decodeURIComponent(atob(str).split('').map(function(c) {{
                return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
            }}).join(''));
        }}
        document.getElementById('copy-btn').addEventListener('click', function() {{
            const text = decodeB64('{b64_json}');
            navigator.clipboard.writeText(text).then(function() {{
                alert('复制成功！排错日志已成功保存，快去发给 nolinda@126.com 吧~');
            }}, function(err) {{
                const textArea = document.createElement("textarea");
                textArea.value = text;
                textArea.style.position = "fixed";
                document.body.appendChild(textArea);
                textArea.focus();
                textArea.select();
                try {{
                    document.execCommand('copy');
                    alert('复制成功！');
                }} catch (err) {{
                    alert('自动复制失败，请直接手动复制下方框内代码。');
                }}
                document.body.removeChild(textArea);
            }});
        }});
        </script>
        """
        components.html(copy_html, height=45)
        st.code(log_json, language="json")
else:
    st.info("💡 请在上方上传【表A】、【表B】。")

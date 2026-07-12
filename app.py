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
    include_k = custom_time_keywords if custom_time_keywords else ['总', '时长', '时间', '片长', '总长', '总片长', 'Duration', 'time']
    exclude_k = ['上线', '日期', '发布', '开播', '年度', '更新', 'date', 'day', '总集数', '总片数', '总集', 'ID', '号', 'No']
    is_inc = any(k in col_name for k in include_k)
    is_exc = any(k in col_name for k in exclude_k)
    return is_inc and not is_exc

def parse_time_logic(val):
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
    best_idx, max_m = 0, -1
    core = custom_keywords if custom_keywords else {'编码', '名称', '日期', '类型', '状态', '标题', '时间', 'ID', 'Name', 'Date', 'Type'}
    
    for idx in range(min(10, len(df))):
        row = [str(x).strip() for x in df.iloc[idx].values if pd.notna(x)]
        m = sum(3 if any(c in v for c in core) else (0.8 if len(v)>1 else 0) for v in row)
        if len(row) > 0:
            num_count = sum(1 for v in row if v.replace('.','',1).isdigit())
            if num_count / len(row) > 0.5: m -= 5
        if m > max_m: max_m, best_idx = m, idx
        
    h = [str(x).strip() for x in df.iloc[best_idx].values]
    return best_idx, [x for x in h if x and x != 'nan' and not x.startswith('Unnamed:')]

# --- 2. 页面基本配置 ---
st.set_page_config(page_title="万能 Excel 表头提取与排版助手", layout="wide", page_icon="🔀")

# --- 3. 顶部说明区 ---
st.title("🔀 万能 Excel 表头提取与排版助手")

st.markdown("""
✅ 此网页主要解决 Excel 表头/字段对应提取转换（B表字段 and A表字段大部分重合）的问题。

✅ 已实现自动对齐，即 字段/表头字面完全一样、或相似度较高的，系统会自动处理。

✅ 已考虑到个别数据总时长会超过24小时（如长剧集、系列课程、累计工时）的情况。

---

系统会自动把 **【表A（原始数据）】** 里的数据提取/筛选并填入 **【表b（带有你需要格式、字段/表头的模版）】** 里，**最终导出的 Sheet 将完全沿用表b的结构与命名。**

* **自定义填充：** 如果A表没有B表某字段（表B字段比A多），你可以为这些字段设置不同的默认内容，不设置视为空。（详见左侧第 1 点）
* **无表头预警：** 如果A表某列无表头但该列有内容，传完表A后见下方提示，可见左侧第2点手动对号修正；
* **特殊处理：** 只要字段满足时间特征，系统就会默认时、分、秒的无限累计显示（`[h]:mm:ss`）（可取消）；
* **求助与售后 (Help & Support)：** * 🔍 **自查格式**：报错转换失败先自查表b表头在哪一行，见左侧第3点自查—表B表头行号修正。
    * 🔄 **页面卡顿/未刷新**：可点击网页最右上角 **三点菜单 (Three-dot Menu `⋮`)**：
        * 👉 选择 **Rerun (重新运行)** 强制页面刷新重算。
        * 👉 选择 **Clear cache (清除缓存)** 清空之前的表格记忆（上传新模板疯狂报错时必点）。
    * 🧑‍💻 **联系售后**：若尝试上述方法仍未解决，请点击底部【排错专区】的复制按钮，将报错信息发给 **nolinda@126.com**。
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
            if len(parts) == 2:
                custom_defaults[parts[0].strip()] = parts[1].strip()

# --- 5. 侧边栏：2. 手动对号修正 ---
st.sidebar.markdown("---")
st.sidebar.header("🛠️ 2. 手动对号修正")
st.sidebar.info("👉 **用法：** 表B字段名=表A列号。例如：UMAI=58(表b叫umail的列 对应的是表a第58列），输入完按回车键或在空白处点一下即可生效。")
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
with col_u1:
    raw_file = st.file_uploader("📂 上传【表A：数据源】(系统导出的原始表)", type=["csv", "xlsx", "xls"])
with col_u2:
    template_file = st.file_uploader("📋 上传【表B：目标模板】(带有目标格式的模版)", type=["xlsx", "xls"])

# ==================== 7. 读取表A与无头预警 ====================
df_raw = None
raw_sheets = {} 
if raw_file:
    try:
        ext = raw_file.name.split('.')[-1].lower()
        if ext == 'csv':
            df_raw = pd.read_csv(raw_file, dtype=str).fillna("")
            raw_sheets["Sheet1"] = df_raw
        else:
            xls_raw = pd.ExcelFile(raw_file)
            for s in xls_raw.sheet_names:
                raw_sheets[s] = pd.read_excel(xls_raw, sheet_name=s, dtype=str).fillna("")
            df_raw = list(raw_sheets.values())[0] 
            
        st.success(f"✅ 表A 读取成功！共读取 {len(raw_sheets)} 个 Sheet。")
        
        unnamed = []
        for s_name, df_s in raw_sheets.items():
            for i, col in enumerate(df_s.columns, 1):
                if "Unnamed" in str(col) or str(col).strip() == "":
                    sample = [str(x).strip() for x in df_s.iloc[:, i-1].tolist() if str(x).strip() != ""]
                    unnamed.append((s_name, i, "、".join(sample[:3]) if sample else "全空"))
        
        if unnamed:
            st.warning("🚨 **表A 发现无名列！** 请根据内容预览，在左侧【手动对号修正】区指定列号：")
            for s_name, idx, pre in unnamed: 
                sheet_prefix = f"『{s_name}』 " if len(raw_sheets) > 1 else ""
                st.write(f"👉 表A {sheet_prefix}第 `{idx}` 列 ➔ 内容预览: `{pre}`")
    except Exception as e: 
        st.error(f"读取表A失败: {e}")
        st.stop()

# ==================== 8. 侧边栏：3. 多Sheet指派与特征词 ====================
selected_time_cols = []
force_header_config = {}
sheet_mapping = {}

st.sidebar.markdown("---")
st.sidebar.header("🔍 3. 数据流向与表头行自查")

header_keywords_input = st.sidebar.text_input(
    "🧠 表头定位特征词 (逗号隔开)",
    value="编码,名称,日期,类型,状态,标题,时间,片单,导演,演员,ID,Name,Date,Type",
    help="辅助 AI 定位表头在哪一行。"
)
custom_header_keywords = set(x.strip() for x in header_keywords_input.replace("，", ",").split(",") if x.strip())

time_keywords_input = st.sidebar.text_input(
    "⏳ 时间列特征词 (逗号隔开)",
    value="总,时长,时间,片长,总长,总片长,Duration,time"
)
custom_time_keywords = [x.strip() for x in time_keywords_input.replace("，", ",").split(",") if x.strip()]

if template_file and raw_sheets:
    xls_tpl = pd.ExcelFile(template_file)
    a_sheet_names = list(raw_sheets.keys())
    
    st.sidebar.info("👇 **请为表B的每个Sheet指定数据源及表头位置：**")
    
    for s_name in xls_tpl.sheet_names:
        with st.sidebar.expander(f"📁 表B Sheet: 『{s_name}』", expanded=True):
            # 下拉框：指派表A的哪个Sheet
            sheet_mapping[s_name] = st.selectbox(
                "该Sheet提取自哪里的数据？",
                options=a_sheet_names,
                key=f"map_{s_name}"
            )
            # 行号强校验
            force_header_config[s_name] = st.number_input(
                f"表头在表B第几行？(0为自动侦测)", 
                min_value=0, max_value=20, value=0, 
                key=f"row_{s_name}"
            )
    
    # 时间列统一确认
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
elif template_file and not raw_sheets:
    st.sidebar.warning("请先上传表A，再进行指派。")

# ==================== 9. 执行核心映射 ====================
if df_raw is not None and template_file:
    output = BytesIO()
    final_reports = {}
    debug_log = {"诊断": {}}

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for s_name in xls_tpl.sheet_names:
            df_tpl_meta = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=10)
            if df_tpl_meta.empty: continue
            
            # 精准获取下拉框指派的表A Sheet
            target_a_sheet_name = sheet_mapping.get(s_name)
            if not target_a_sheet_name or target_a_sheet_name not in raw_sheets:
                continue # 如果没选或选错，跳过
                
            current_df_raw = raw_sheets[target_a_sheet_name]
            if current_df_raw.empty:
                st.warning(f"⚠️ 警告：指派给『{s_name}』的表A Sheet（{target_a_sheet_name}）内无数据！")
            
            f_h = force_header_config.get(s_name, 0)
            h_idx, headers = (f_h-1, [str(x).strip() for x in df_tpl_meta.iloc[f_h-1].values if pd.notna(x)]) if f_h > 0 else find_headers(df_tpl_meta, custom_header_keywords)
            
            out_df = pd.DataFrame(columns=headers)
            raw_cols = current_df_raw.columns.tolist()
            report = []
            
            for b_idx, col_name in enumerate(headers, 1):
                a_idx, status = None, "empty"
                
                if col_name.lower() in manual_map_config:
                    a_idx = manual_map_config[col_name.lower()]; status = "ok"
                else:
                    m = difflib.get_close_matches(col_name, raw_cols, n=1, cutoff=0.4)
                    if m: a_idx = raw_cols.index(m[0]); status = "ok"
                
                if status == "ok" and a_idx < len(current_df_raw.columns):
                    series = current_df_raw.iloc[:, a_idx]
                    if col_name in selected_time_cols:
                        series = series.apply(parse_time_logic)
                    out_df[col_name] = series
                    report.append({"b": b_idx, "bn": col_name, "a": a_idx + 1, "an": raw_cols[a_idx], "s": "ok"})
                else:
                    fill = custom_defaults.get(col_name, "")
                    out_df[col_name] = fill
                    report.append({"b": b_idx, "bn": col_name, "f": fill, "s": "fill" if fill else "empty"})
            
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
            debug_log["诊断"][s_name] = {"数据源": target_a_sheet_name, "识别行": h_idx+1, "表B字段": headers}

    st.markdown("---")
    st.subheader("📊 数据对齐与填充看板 (表A ➡️ 表B)")

    for s_name, report in final_reports.items():
        ok_count = sum(1 for x in report if x['s'] == 'ok')
        with st.expander(f"📁 『{s_name}』 预览 (提取成功: {ok_count} | 填充/留空: {len(report)-ok_count})"):
            c1, c2 = st.columns(2)
            with c1:
                st.write("🟢 **a表提取成功**")
                for x in [item for item in report if item['s'] == 'ok']:
                    st.write(f"A第`{x['a']}`列({x['an']}) ➔ B第`{x['b']}`列({x['bn']})")
                if not any(item['s'] == 'ok' for item in report):
                    st.caption("没有成功从表A提取的数据。")
            with c2:
                st.write("🟡 **b表填充或空**")
                for x in [item for item in report if item['s'] != 'ok']:
                    val_text = f"填: `{x['f']}`" if x['f'] else "留空"
                    st.write(f"B第`{x['b']}`列({x['bn']}) ➔ {val_text}")

    st.success("🎉 处理完成！")
    st.download_button("📥 下载转换后的交付 Excel", data=output.getvalue(), file_name=f"交付表_{today_str}.xlsx")
    
    # --- 10. 排错专区诊断日志与一键复制组件 ---
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
                // 应急方案
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

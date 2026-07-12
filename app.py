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
    """
    智能识别时间/时长列：
    排除包含日期、上线、更新等不适合累加的干扰词
    """
    include_k = custom_time_keywords if custom_time_keywords else ['总', '时长', '时间', '片长', '总长', '总片长', 'Duration', 'time']
    exclude_k = ['上线', '日期', '发布', '开播', '年度', '更新', 'date', 'day', '总集数', '总片数', '总集', 'ID', '号', 'No']
    is_inc = any(k in str(col_name) for k in include_k)
    is_exc = any(k in str(col_name) for k in exclude_k)
    return is_inc and not is_exc

def parse_time_logic(val):
    """支持超过 24 小时的时间数据强制解析"""
    if pd.isna(val) or val == "": return ""
    v_str = str(val).strip()
    if ':' in v_str:
        parts = v_str.split(':')
        try:
            if len(parts) == 3: # HH:MM:SS 格式
                return (int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])) / 86400.0
            elif len(parts) == 2: # MM:SS 格式
                return (int(parts[0]) * 60 + int(parts[1])) / 84600.0
        except: pass
    try: return float(v_str)
    except: return v_str

def find_headers(df, custom_keywords=None):
    """
    智能侦测表头所在行：
    1. 扩大扫描至前200行（应对极其变态的超长表头说明区）
    2. 根据特征词命中率打分
    3. 排除纯数字数据行干扰
    """
    best_idx, max_m = 0, -1
    core = custom_keywords if custom_keywords else {'编码', '名称', '日期', '类型', '状态', '标题', '时间', 'ID', 'Name', 'Date', 'Type'}
    
    # 扩大雷达：最多扫描前 200 行
    for idx in range(min(200, len(df))):
        row = [str(x).strip() for x in df.iloc[idx].values if pd.notna(x)]
        m = sum(3 if any(c in v for c in core) else (0.8 if len(v)>1 else 0) for v in row)
        
        if len(row) > 0:
            num_count = sum(1 for v in row if v.replace('.','',1).isdigit())
            if num_count / len(row) > 0.5: m -= 5
                
        if m > max_m: max_m, best_idx = m, idx
        
    # 保留表头原貌，如果是空表头则赋予占位符，防止错位
    h_row = df.iloc[best_idx].values
    clean_h = [str(x).strip() if pd.notna(x) and str(x).strip() != '' else f"Unnamed_{i}" for i, x in enumerate(h_row)]
    return best_idx, clean_h

# --- 2. 页面基本配置 ---
st.set_page_config(page_title="万能 Excel 表头提取与排版助手", layout="wide", page_icon="🔀")

# --- 3. 顶部说明区 ---
st.title("🔀 万能 Excel 表头提取与排版助手")

st.markdown("""
✅ 此网页主要解决 Excel 表头/字段对应提取转换（B表字段 and A表字段大部分重合）的问题。
✅ 已实现自动对齐，即字段/表头字面完全一样、或相似度较高的，系统会自动处理。
✅ 已考虑到个别数据总时长会超过24小时（如长剧集、系列课程、累计工时）的情况。
---
系统会自动把 **【表A（原始数据）】** 里的数据提取/筛选并填入 **【表B（带有你需要格式、字段的模版）】** 里，**最终导出的 Sheet 将完全沿用表B的结构与命名。**

* **自定义填充：** 如果A表没有B表某字段，你可以为这些字段设置默认内容，不设置视为空。（详见左侧第 1 点）
* **无表头预警：** 如果A表某列无表头但该列有内容，传完表A后见下方提示，可见左侧第 2 点手动对号修正。
* **特殊处理：** 只要字段满足时间特征，系统会默认时、分、秒的无限累计显示（`[h]:mm:ss`）。
* **求助与售后：** 若转换失败或报错，请点击底部【排错专区】一键复制日志，发给 **nolinda@126.com**。
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
st.sidebar.info("👉 **用法：** 表B字段名=表A列号。例如：UMAI=58，输入完按回车生效。")
manual_map_text = st.sidebar.text_area("输入映射规则（不区分大小写）：", placeholder="许可证=10", height=120)

manual_map_config = {}
if manual_map_text.strip():
    for line in manual_map_text.split('\n'):
        if '=' in line:
            parts = line.split('=')
            if len(parts) == 2:
                try: 
                    manual_map_config[parts[0].strip().lower()] = int(parts[1].strip()) - 1
                except ValueError: pass

# --- 6. 主界面：文件上传 ---
col_u1, col_u2 = st.columns(2)
with col_u1:
    raw_file = st.file_uploader("📂 上传【表A：数据源】(系统导出的原始表)", type=["csv", "xlsx", "xls"])
with col_u2:
    template_file = st.file_uploader("📋 上传【表B：目标模板】(带有目标格式的模版)", type=["xlsx", "xls"])

# ==================== 7. 读取表A与无头预警 ====================
df_raw = None
if raw_file:
    try:
        ext = raw_file.name.split('.')[-1].lower()
        # 智能侦测表 A 的真实表头行
        if ext == 'csv':
            temp_a = pd.read_csv(raw_file, header=None, nrows=50, dtype=str).fillna("")
            raw_file.seek(0)
        else:
            temp_a = pd.read_excel(raw_file, header=None, nrows=50, dtype=str).fillna("")
            raw_file.seek(0)
            
        a_idx, _ = find_headers(temp_a, {'编码', '名称', '日期', '类型', '状态', '标题', '时间', 'ID'})
        
        # 使用真实的表头行读取表 A
        if ext == 'csv':
            df_raw = pd.read_csv(raw_file, header=a_idx, dtype=str).fillna("")
        else:
            df_raw = pd.read_excel(raw_file, header=a_idx, dtype=str).fillna("")
            
        st.success(f"✅ 表A 读取成功！(识别到表头在第 {a_idx + 1} 行) 共包含 {len(df_raw.columns)} 列数据。")
        
        # 提取无名列预览
        unnamed = []
        for i, col in enumerate(df_raw.columns, 1):
            if "Unnamed" in str(col) or str(col).strip() == "":
                sample = [str(x).strip() for x in df_raw.iloc[:, i-1].tolist() if str(x).strip() != ""]
                unnamed.append((i, "、".join(sample[:3]) if sample else "全空"))
        
        if unnamed:
            st.warning("🚨 **表A 发现无名列！** 请根据内容预览，在左侧【手动对号修正】区指定列号：")
            for idx, pre in unnamed: 
                st.write(f"👉 表A 第 `{idx}` 列 ➔ 内容预览: `{pre}`")
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
    help="系统会根据这些核心词自动定位最像表头的那一行。支持任何行业词汇。"
)
custom_header_keywords = set(x.strip() for x in header_keywords_input.replace("，", ",").split(",") if x.strip())

time_keywords_input = st.sidebar.text_input(
    "⏳ 自定义时间列特征词 (逗号隔开)",
    value="总,时长,时间,片长,总长,总片长,Duration,time",
    help="表头只要包含这些词，就会被默认勾选进入 [h]:mm:ss 的无限时间累加转换。"
)
custom_time_keywords = [x.strip() for x in time_keywords_input.replace("，", ",").split(",") if x.strip()]

if template_file:
    xls_tpl = pd.ExcelFile(template_file)
    
    # 动态渲染 Sheet 的表头行号输入（表B第几行）
    for s_name in xls_tpl.sheet_names:
        force_header_config[s_name] = st.sidebar.number_input(
            f"『{s_name}』的字段名在表B第几行？(0为自动识别)", 
            min_value=0, max_value=200, value=0, 
            key=f"row_{s_name}"
        )
    
    # ⏳ 时间格式转换确认
    st.sidebar.markdown("---")
    st.sidebar.subheader("⏳ 时间格式转换确认")
    
    all_time_cols_in_template = []
    for s_name in xls_tpl.sheet_names:
        try:
            # 扩大扫描范围到 200 行，提取时间列特征
            df_preview = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=200)
            if not df_preview.empty:
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
    st.sidebar.caption("⏳ 上传【表B】后，即可在此进行行号自查与时间列转换设置。")

# ==================== 9. 执行核心映射与渲染看板 ====================
if df_raw is not None and template_file:
    output = BytesIO()
    final_reports = {}
    debug_log = {"源表字段": df_raw.columns.tolist(), "诊断": {}}

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for s_name in xls_tpl.sheet_names:
            # 必须读取足够多的行，防止表头极其靠下
            df_tpl_meta = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=200)
            if df_tpl_meta.empty: continue
            
            f_h = force_header_config.get(s_name, 0)
            if f_h > 0:
                h_idx = f_h - 1
                headers = [str(x).strip() if pd.notna(x) else f"Unnamed_{i}" for i, x in enumerate(df_tpl_meta.iloc[h_idx].values)]
            else:
                h_idx, headers = find_headers(df_tpl_meta, custom_header_keywords)
            
            out_df = pd.DataFrame(columns=headers)
            raw_cols = df_raw.columns.tolist()
            report = []
            
            for b_idx, col_name in enumerate(headers, 1):
                a_idx, status = None, "empty"
                
                # 过滤掉系统为了防错生成的占位符名称
                if col_name.startswith("Unnamed_"):
                    out_df[col_name] = ""
                    continue
                
                # 优先级1：手动列号映射（忽略大小写）
                if col_name.lower() in manual_map_config:
                    a_idx = manual_map_config[col_name.lower()]; status = "ok"
                # 优先级2：自动相似度对齐
                else:
                    m = difflib.get_close_matches(col_name, raw_cols, n=1, cutoff=0.4)
                    if m: a_idx = raw_cols.index(m[0]); status = "ok"
                
                if status == "ok" and a_idx < len(df_raw.columns):
                    series = df_raw.iloc[:, a_idx]
                    if col_name in selected_time_cols:
                        series = series.apply(parse_time_logic)
                    out_df[col_name] = series
                    report.append({"b": b_idx, "bn": col_name, "a": a_idx + 1, "an": raw_cols[a_idx], "s": "ok"})
                else:
                    fill = custom_defaults.get(col_name, "")
                    out_df[col_name] = fill
                    report.append({"b": b_idx, "bn": col_name, "f": fill, "s": "fill" if fill else "empty"})
            
            # 写入 sheet，去除多余的列标题行
            out_df.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # 格式化时间列并还原真实的空表头显示
            for i, h in enumerate(headers, 1):
                if h.startswith("Unnamed_"):
                    ws.cell(row=1, column=i).value = ""  # 擦除占位符
                elif h in selected_time_cols:
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 14
                    for r in range(2, len(out_df)+2):
                        cell = ws.cell(row=r, column=i)
                        if isinstance(cell.value, (float, int)): 
                            cell.number_format = '[h]:mm:ss'
            
            final_reports[s_name] = report
            debug_log["诊断"][s_name] = {"识别行": h_idx+1, "字段": headers}

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

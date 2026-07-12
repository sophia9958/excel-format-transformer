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
    1. 根据特征词命中率打分
    2. 评估文本长度和密度
    3. 排除大部分是纯数字的干扰行
    """
    best_idx, max_m = 0, -1
    core = custom_keywords if custom_keywords else {'编码', '名称', '日期', '类型', '状态', '标题', '时间', 'ID', 'Name', 'Date', 'Type'}
    
    for idx in range(min(10, len(df))):
        row = [str(x).strip() for x in df.iloc[idx].values if pd.notna(x)]
        # 基础分：含有特征词权重极高，包含长度大于1的文本有基础分
        m = sum(3 if any(c in v for c in core) else (0.8 if len(v)>1 else 0) for v in row)
        
        # 降噪：如果整行有一半以上是纯数字，说明大概率是数据行而非表头，扣除大分
        if len(row) > 0:
            num_count = sum(1 for v in row if v.replace('.','',1).isdigit())
            if num_count / len(row) > 0.5:
                m -= 5
                
        if m > max_m: max_m, best_idx = m, idx
        
    h = [str(x).strip() for x in df.iloc[best_idx].values]
    return best_idx, [x for x in h if x and x != 'nan' and not x.startswith('Unnamed:')]

def normalize_text(value):
    """用于跨 Sheet 比较的统一文本格式。"""
    if pd.isna(value):
        return ""
    return str(value).strip().lower().replace(" ", "")

def read_excel_sheets(excel_file):
    """读取 Excel 中所有非空 Sheet；返回 {sheet名: DataFrame}。"""
    xls = pd.ExcelFile(excel_file)
    result = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str).fillna("")
        if not df.empty:
            result[sheet_name] = df
    return result

def build_template_profiles(xls_tpl, force_header_config, custom_header_keywords):
    """读取表B每个 Sheet 的字段、原有数据和可用于分流的特征值。"""
    profiles = {}
    for sheet_name in xls_tpl.sheet_names:
        meta = pd.read_excel(xls_tpl, sheet_name=sheet_name, header=None, nrows=10)
        if meta.empty:
            continue
        forced = force_header_config.get(sheet_name, 0)
        if forced > 0:
            header_idx = forced - 1
            headers = [str(x).strip() for x in meta.iloc[header_idx].values
                       if pd.notna(x) and str(x).strip()]
        else:
            header_idx, headers = find_headers(meta, custom_header_keywords)
        body = pd.read_excel(xls_tpl, sheet_name=sheet_name, header=header_idx,
                             dtype=str).fillna("")
        profiles[sheet_name] = {
            "header_idx": header_idx,
            "headers": headers,
            "body": body,
        }
    return profiles

def route_rows_to_template_sheets(df_raw, template_profiles, routing_rules=None):
    """
    把表A每一行只分配给表B的一个 Sheet。
    先用编码/ID/名称等唯一字段精确匹配；模板尚无该条记录时，
    再使用用户在页面上为各 Sheet 选择的通用字段和值进行筛选。
    """
    sheet_names = list(template_profiles)
    routed_indices = {name: [] for name in sheet_names}
    if len(sheet_names) <= 1:
        if sheet_names:
            routed_indices[sheet_names[0]] = df_raw.index.tolist()
        return routed_indices, []

    raw_cols = list(df_raw.columns)
    identity_words = ("编码", "代码", "id", "编号", "名称", "标题", "name")
    routing_rules = routing_rules or {}

    # 提前缓存各 Sheet 每个共同字段中出现过的值，避免逐行重复计算。
    value_sets = {}
    for sheet_name, profile in template_profiles.items():
        body = profile["body"]
        value_sets[sheet_name] = {
            col: {normalize_text(v) for v in body[col].tolist() if normalize_text(v)}
            for col in raw_cols if col in body.columns
        }

    unresolved = []
    for idx, row in df_raw.iterrows():
        exact_hits = []
        for sheet_name in sheet_names:
            for col, values in value_sets[sheet_name].items():
                col_lower = str(col).lower()
                value = normalize_text(row[col])
                if value and any(word in col_lower for word in identity_words) and value in values:
                    exact_hits.append(sheet_name)
                    break

        # 唯一字段只命中一个 Sheet 时，归属最可靠。
        unique_hits = list(dict.fromkeys(exact_hits))
        if len(unique_hits) == 1:
            routed_indices[unique_hits[0]].append(idx)
            continue

        rule_hits = []
        for sheet_name in sheet_names:
            rules = routing_rules.get(sheet_name, {})
            # 该 Sheet 至少要配置一个字段；多个字段是 AND，同字段多值是 OR。
            active_rules = {col: values for col, values in rules.items() if values}
            if active_rules and all(
                col in row.index and normalize_text(row[col]) in values
                for col, values in active_rules.items()
            ):
                rule_hits.append(sheet_name)

        if len(rule_hits) == 1:
            routed_indices[rule_hits[0]].append(idx)
        else:
            unresolved.append(idx)

    return routed_indices, unresolved

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

# 解析多字段默认值
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

# 解析手动映射规则 (忽略大小写)
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
raw_sheet_names = []
if raw_file:
    try:
        ext = raw_file.name.split('.')[-1].lower()
        if ext == 'csv':
            df_raw = pd.read_csv(raw_file, dtype=str).fillna("")
            raw_sheet_names = [raw_file.name]
        else:
            raw_sheets = read_excel_sheets(raw_file)
            raw_sheet_names = list(raw_sheets)
            # 表A有多个 Sheet 时统一参与分流；同名字段自动纵向合并。
            df_raw = pd.concat(raw_sheets.values(), ignore_index=True, sort=False).fillna("")
        st.success(
            f"✅ 表A 读取成功！共 {len(df_raw)} 条、{len(df_raw.columns)} 列"
            f"，有效 Sheet：{len(raw_sheet_names)} 个。"
        )
        
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
routing_rules = {}
template_profiles = {}

st.sidebar.markdown("---")
st.sidebar.header("🔍 3. 自查—表B表头行号修正")

# 让用户可以完全动态自定义表头特征词
header_keywords_input = st.sidebar.text_input(
    "🧠 自定义表头定位特征词 (逗号隔开)",
    value="编码,名称,日期,类型,状态,标题,时间,片单,导演,演员,ID,Name,Date,Type",
    help="系统会根据这些核心词在表格前10行中自动定位最像表头的那一行。支持任何行业词汇。"
)
custom_header_keywords = set(x.strip() for x in header_keywords_input.replace("，", ",").split(",") if x.strip())

# 让用户可以完全动态自定义时间列特征词
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
            f"『{s_name}』的字段名在表B第几行？", 
            min_value=0, max_value=20, value=0, 
            key=f"row_{s_name}"
        )
    
    # ⏳ 时间格式转换确认 (紧随在第3点自查后面)
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

    # 🔀 完全通用的分 Sheet 规则：不内置任何行业或业务词。
    if df_raw is not None and len(xls_tpl.sheet_names) > 1:
        template_profiles = build_template_profiles(
            xls_tpl, force_header_config, custom_header_keywords
        )
        common_route_cols = [
            col for col in df_raw.columns
            if any(col in p["body"].columns for p in template_profiles.values())
        ]
        st.sidebar.markdown("---")
        st.sidebar.header("🧭 4. 多 Sheet 分流规则")
        st.sidebar.info(
            "先选择用于分 Sheet 的字段，再给每个 Sheet 勾选允许值。"
            "同一 Sheet 的多个字段必须同时满足；同一字段勾选多个值时满足任一即可。"
            "已存在于表B的编码/名称仍会优先自动归属。"
        )
        route_fields = st.sidebar.multiselect(
            "选择分 Sheet 字段",
            common_route_cols,
            default=[],
            help="优先选择分类、部门、地区、渠道、状态等取值较少且能区分 Sheet 的字段。"
        )

        for s_name in xls_tpl.sheet_names:
            routing_rules[s_name] = {}
            with st.sidebar.expander(f"📁 『{s_name}』允许值", expanded=False):
                for field in route_fields:
                    source_values = sorted({
                        str(v).strip() for v in df_raw[field].tolist()
                        if str(v).strip()
                    })
                    body = template_profiles.get(s_name, {}).get("body", pd.DataFrame())
                    template_values = set()
                    if field in body.columns:
                        template_values = {
                            normalize_text(v) for v in body[field].tolist()
                            if normalize_text(v)
                        }
                    # 默认勾选该 Sheet 模板中已经出现过、且表A也存在的值。
                    suggested = [v for v in source_values if normalize_text(v) in template_values]
                    chosen = st.multiselect(
                        f"{field}", source_values, default=suggested,
                        key=f"route_{s_name}_{field}"
                    )
                    routing_rules[s_name][field] = {
                        normalize_text(v) for v in chosen
                    }
else:
    st.sidebar.caption("⏳ 上传【表B】后，即可在此进行行号自查与时间列转换设置。")

# ==================== 9. 执行核心映射与渲染看板 ====================
if df_raw is not None and template_file:
    output = BytesIO()
    final_reports = {}
    debug_log = {"表A有效Sheet": raw_sheet_names, "源表字段": df_raw.columns.tolist(), "诊断": {}}

    # 关键修复：在写出前先把表A的每一行分配到表B唯一的目标 Sheet。
    if not template_profiles:
        template_profiles = build_template_profiles(
            xls_tpl, force_header_config, custom_header_keywords
        )
    routed_indices, unresolved_indices = route_rows_to_template_sheets(
        df_raw, template_profiles, routing_rules
    )

    if unresolved_indices:
        st.warning(
            f"⚠️ 有 {len(unresolved_indices)} 条记录无法唯一判断目标 Sheet，"
            "为避免乱跑或重复，已不写入任何 Sheet；可在排错日志查看。"
        )
        debug_log["未能分Sheet的表A行号"] = [int(i) + 2 for i in unresolved_indices]

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for s_name in xls_tpl.sheet_names:
            if s_name not in template_profiles: continue
            h_idx = template_profiles[s_name]["header_idx"]
            headers = template_profiles[s_name]["headers"]
            # 每个 Sheet 只使用已经分配给自己的表A记录，不再重复使用整张 df_raw。
            sheet_raw = df_raw.loc[routed_indices.get(s_name, [])].reset_index(drop=True)
            
            out_df = pd.DataFrame(columns=headers)
            raw_cols = sheet_raw.columns.tolist()
            report = []
            
            for b_idx, col_name in enumerate(headers, 1):
                a_idx, status = None, "empty"
                
                # 优先级1：手动列号映射（忽略大小写）
                if col_name.lower() in manual_map_config:
                    a_idx = manual_map_config[col_name.lower()]; status = "ok"
                # 优先级2：自动相似度对齐
                else:
                    m = difflib.get_close_matches(col_name, raw_cols, n=1, cutoff=0.4)
                    if m: a_idx = raw_cols.index(m[0]); status = "ok"
                
                if status == "ok" and a_idx < len(df_raw.columns):
                    series = sheet_raw.iloc[:, a_idx]
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
            debug_log["诊断"][s_name] = {
                "识别行": h_idx+1,
                "字段": headers,
                "写入记录数": len(sheet_raw),
                "表A原始行号": [int(i) + 2 for i in routed_indices.get(s_name, [])]
            }

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
        # 将诊断日志转换为格式化JSON与加密Base64字符串
        log_json = json.dumps(debug_log, ensure_ascii=False, indent=2)
        b64_json = base64.b64encode(log_json.encode('utf-8')).decode('utf-8')
        
        # 页面注入带有JS监听和优雅动画的复制按钮
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
    # 没传文件时的默认占位显示
    st.info("💡 请在上方上传【表A】、【表B】。")

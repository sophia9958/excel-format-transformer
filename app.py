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
    is_inc = any(k in str(col_name) for k in include_k)
    is_exc = any(k in str(col_name) for k in exclude_k)
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
    
    for idx in range(min(15, len(df))):
        row = [str(x).strip() for x in df.iloc[idx].values if pd.notna(x)]
        m = sum(3 if any(c in v for c in core) else (0.8 if len(v)>1 else 0) for v in row)
        if len(row) > 0:
            num_count = sum(1 for v in row if v.replace('.','',1).isdigit())
            if num_count / len(row) > 0.6: m -= 5
        if m > max_m: max_m, best_idx = m, idx
    
    h_row = df.iloc[best_idx].values
    clean_h = [str(x).strip() if pd.notna(x) else f"Unnamed: {i}" for i, x in enumerate(h_row)]
    return best_idx, clean_h

# --- 2. 页面配置 ---
st.set_page_config(page_title="万能 Excel 表头提取助手", layout="wide", page_icon="🔀")

# --- 3. 顶部说明 ---
st.title("🔀 万能 Excel 表头提取与排版助手")
st.markdown("""
✅ 解决 B 表字段与 A 表字段对应提取转换问题，支持相似度自动对齐。
✅ **[h]:mm:ss** 累计时长自动转换，无视 24 小时进位限制。
✅ **多 Sheet 智能寻呼**：表 A 与表 B 同名 Sheet 自动对应读取。
---
📬 **联系售后**：若遇报错，请复制报错信息发送至 **nolinda@126.com**。
""")

# --- 4. 侧边栏配置 ---
st.sidebar.header("⚙️ 参数配置")
header_keywords_input = st.sidebar.text_input("🧠 表头定位特征词", value="编码,名称,日期,类型,状态,标题,时间,ID,片单")
custom_header_keywords = set(x.strip() for x in header_keywords_input.replace("，", ",").split(",") if x.strip())

time_keywords_input = st.sidebar.text_input("⏳ 时长识别特征词", value="总,时长,时间,片长,总长,Duration")
custom_time_keywords = [x.strip() for x in time_keywords_input.replace("，", ",").split(",") if x.strip()]

st.sidebar.markdown("---")
st.sidebar.header("🛠️ 手动修正与默认填充")
custom_defaults_text = st.sidebar.text_area("1. 默认填充 (字段=内容)", placeholder="更新日期=260511", height=100)
manual_map_text = st.sidebar.text_area("2. 手动对号 (B字段=A列号)", placeholder="许可证=10", height=100)

custom_defaults = {l.split('=')[0].strip(): l.split('=')[1].strip() for l in custom_defaults_text.split('\n') if '=' in l}
manual_map_config = {l.split('=')[0].strip().lower(): int(l.split('=')[1].strip())-1 for l in manual_map_text.split('\n') if '=' in l and l.split('=')[1].strip().isdigit()}

# --- 5. 文件上传与逻辑处理 ---
u1, u2 = st.columns(2)
with u1: raw_file = st.file_uploader("📂 上传【表 A：源数据】", type=["csv", "xlsx", "xls"])
with u2: template_file = st.file_uploader("📋 上传【表 B：目标模板】", type=["xlsx", "xls"])

if raw_file and template_file:
    try:
        output = BytesIO()
        xls_tpl = pd.ExcelFile(template_file)
        
        # 判断表 A 是 Excel 还是单薄的 CSV
        is_raw_excel = not raw_file.name.endswith('.csv')
        xls_raw = pd.ExcelFile(raw_file) if is_raw_excel else None
        raw_sheet_names = xls_raw.sheet_names if is_raw_excel else []

        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for s_name in xls_tpl.sheet_names:
                # 核心升级：为表 B 的当前 Sheet 寻找表 A 的对应 Sheet
                if is_raw_excel:
                    # 如果表 A 也有同名 Sheet，就读取同名；否则默认读第一个
                    target_raw_sheet = s_name if s_name in raw_sheet_names else raw_sheet_names[0]
                    temp_df = pd.read_excel(xls_raw, sheet_name=target_raw_sheet, header=None, nrows=20, dtype=str).fillna("")
                else:
                    raw_file.seek(0)
                    temp_df = pd.read_csv(raw_file, header=None, nrows=20, dtype=str).fillna("")

                # 寻找表 A 表头
                a_h_idx, _ = find_headers(temp_df, custom_header_keywords)
                
                # 正式读取表 A 的当前工作表数据
                if is_raw_excel:
                    df_raw_current = pd.read_excel(xls_raw, sheet_name=target_raw_sheet, header=a_h_idx, dtype=str).fillna("")
                else:
                    raw_file.seek(0)
                    df_raw_current = pd.read_csv(raw_file, header=a_h_idx, dtype=str).fillna("")
                
                # 预警系统：检测当前使用的数据表是否有无名列
                unnamed = []
                for i, col in enumerate(df_raw_current.columns, 1):
                    col_str = str(col).strip()
                    if "Unnamed" in col_str or col_str == "" or col_str.startswith("Column"):
                        sample = [str(x).strip() for x in df_raw_current.iloc[:, i-1].tolist() if str(x).strip() != ""]
                        unnamed.append((i, "、".join(sample[:3]) if sample else "全空"))
                
                if unnamed:
                    st.warning(f"🚨 **表 A ({target_raw_sheet if is_raw_excel else 'CSV文件'}) 发现无名列！** 请在左侧修正：")
                    for idx, pre in unnamed: st.write(f"👉 第 `{idx}` 列 ➔ 内容预览: `{pre}`")

                # 开始处理模板的当前 Sheet
                df_tpl_meta = pd.read_excel(xls_tpl, sheet_name=s_name, header=None, nrows=15).fillna("")
                h_idx, headers = find_headers(df_tpl_meta, custom_header_keywords)
                
                out_df = pd.DataFrame(columns=headers)
                raw_cols = df_raw_current.columns.tolist()
                time_cols_to_format = [h for h in headers if is_duration_col(h, custom_time_keywords)]
                
                for col_name in headers:
                    a_idx = manual_map_config.get(col_name.lower())
                    if a_idx is None:
                        m = difflib.get_close_matches(col_name, raw_cols, n=1, cutoff=0.4)
                        if m: a_idx = raw_cols.index(m[0])
                    
                    if a_idx is not None and a_idx < len(df_raw_current.columns):
                        series = df_raw_current.iloc[:, a_idx]
                        if col_name in time_cols_to_format: series = series.apply(parse_time_logic)
                        out_df[col_name] = series
                    else:
                        out_df[col_name] = custom_defaults.get(col_name, "")
                
                out_df.to_excel(writer, sheet_name=s_name, index=False)
                
                # 时间格式渲染
                ws = writer.sheets[s_name]
                for i, h in enumerate(headers, 1):
                    if h in time_cols_to_format:
                        for r in range(2, len(out_df)+2):
                            cell = ws.cell(row=r, column=i)
                            if isinstance(cell.value, (float, int)): cell.number_format = '[h]:mm:ss'

        st.success("🎉 多 Sheet 数据匹配完成！数据已精准入座！")
        st.download_button("📥 下载交付表", data=output.getvalue(), file_name=f"Result_{datetime.datetime.now().strftime('%y%m%d')}.xlsx")
        
    except Exception as e:
        st.error(f"处理失败，请截图发送邮箱: {e}")

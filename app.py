import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import time
import base64

# --- 1. 页面配置与 CSS ---
st.set_page_config(page_title="小雷Excel批量助手", page_icon="⚡", layout="wide")

st.markdown("""
    <style>
    /* 全局背景与字体 */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    
    .main {
        background-color: #fcfdfe;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    }
    
    .block-container { 
        padding-top: 2rem !important; 
        max-width: 1000px !important;
    }
    
    header {visibility: hidden;}
    footer {visibility: hidden;}

    /* 标题样式 */
    .main-title {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.5rem;
        text-align: center;
        background: linear-gradient(90deg, #3b82f6, #06b6d4);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .sub-title {
        font-size: 0.9rem;
        color: #64748b;
        text-align: center;
        margin-bottom: 2rem;
    }

    /* 卡片通用样式 */
    .stFileUploader section {
        border: 2px dashed #e2e8f0 !important;
        background-color: #ffffff !important;
        padding: 1.5rem !important;
        border-radius: 12px !important;
        transition: all 0.3s ease;
    }
    .stFileUploader section:hover {
        border-color: #3b82f6 !important;
        background-color: #f8fafc !important;
    }

    .upload-card {
        background: #ffffff;
        border: 1px solid #f1f5f9;
        padding: 20px;
        border-radius: 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        margin-bottom: 20px;
    }

    .config-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 18px;
        border-radius: 12px;
        margin-bottom: 15px;
    }

    .personal-box {
        padding: 15px;
        background: #ffffff;
        border-radius: 12px;
        margin-top: 10px;
        border: 1px solid #f1f5f9;
        box-shadow: inset 0 2px 4px 0 rgba(0, 0, 0, 0.02);
        border-left: 4px solid #3b82f6;
    }

    .download-bar {
        background: #ecfdf5;
        border: 1px solid #10b981;
        padding: 15px 20px;
        border-radius: 12px;
        margin: 20px 0;
        display: flex;
        align-items: center;
        justify-content: space-between;
        color: #065f46;
        font-weight: 500;
    }

    /* 按钮美化 */
    .stButton > button {
        border-radius: 10px !important;
        padding: 0.5rem 1rem !important;
        font-weight: 600 !important;
        transition: all 0.2s ease !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.25);
    }
    
    /* 侧边栏/输入框样式 */
    h5 {
        margin-bottom: 0.8rem !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        color: #1e293b !important;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .stExpander {
        border-radius: 12px !important;
        border: 1px solid #f1f5f9 !important;
        background: #ffffff !important;
    }
    </style>
    """, unsafe_allow_html=True)

if 'batch_results' not in st.session_state:
    st.session_state.batch_results = []
if 'extract_results' not in st.session_state:
    st.session_state.extract_results = []

# --- 2. 辅助函数 ---
def to_xlsx_stream(file):
    try:
        file.seek(0)
        df = pd.read_excel(file, header=None, engine='xlrd')
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False)
        output.seek(0)
        return output
    except: return None

def clean_columns(df):
    if df is None: return None
    new_cols, seen = [], {}
    
    for i, col in enumerate(df.columns):
        # 提取非 Unnamed 的部分
        parts = []
        if isinstance(col, tuple):
            parts = [str(p).strip() for p in col if p and "Unnamed" not in str(p)]
        else:
            p = str(col).strip()
            if p and "Unnamed" not in p:
                parts = [p]
        
        # 组合名称，如果为空则设为 未命名
        name = " - ".join(parts) if parts else f"未命名_{i}"
        
        # 处理重名
        if name in seen:
            seen[name] += 1
            final_name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
            final_name = name
            
        new_cols.append(final_name)
    
    # 确保列数一致，不再截断，直接赋值
    df.columns = new_cols
    return df

def load_excel(file, start_row, row_count):
    try:
        file.seek(0)
        curr = to_xlsx_stream(file) if file.name.lower().endswith('.xls') else file
        df = pd.read_excel(curr, header=list(range(start_row, start_row + row_count)), engine='openpyxl')
        return clean_columns(df)
    except: return None

def get_headers_only(file, start_row, row_count):
    try:
        file.seek(0)
        curr = to_xlsx_stream(file) if file.name.lower().endswith('.xls') else file
        # 只读取 0 行数据，仅获取表头结构
        df = pd.read_excel(curr, header=list(range(start_row, start_row + row_count)), nrows=0, engine='openpyxl')
        return clean_columns(df).columns.tolist()
    except: return []

def find_col_index(target, header_list):
    if not target: return -1
    try:
        # 直接匹配扁平化后的全名 (UI 看到什么，这里就匹配什么)
        return header_list.index(target) + 1
    except:
        # 如果全名没匹配上，尝试兼容性匹配 (去掉重名后缀)
        clean_t = target.rsplit('_', 1)[0] if '_' in target else target
        for i, h in enumerate(header_list):
            if h and clean_t in str(h): return i + 1
    return -1

def queue_download_js(results):
    import json
    files_list = []
    for name, data in results:
        b64 = base64.b64encode(data).decode()
        files_list.append({"name": name, "base64": b64})
    
    files_json = json.dumps(files_list)
    
    js_code = f"""
        <script>
        (function() {{
            const files = {files_json};
            async function download() {{
                for (const file of files) {{
                    const blob = await (await fetch(`data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${{file.base64}}`)).blob();
                    const url = window.URL.createObjectURL(blob);
                    const link = document.createElement('a');
                    link.style.display = 'none';
                    link.href = url;
                    link.download = file.name;
                    document.body.appendChild(link);
                    link.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(link);
                    await new Promise(r => setTimeout(r, 800));
                }}
            }}
            download();
        }})();
        </script>
    """
    return js_code

# --- 3. 界面交互 ---
st.markdown('<div class="main-title">⚡ 小雷 Excel 批量助手</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">高效、简单、本地处理的 Excel 办公利器</div>', unsafe_allow_html=True)

tab1, tab2 = st.tabs(["🔄 批量关联更新", "✨ 批量字段提取"])

with tab1:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    up_c1, up_c2 = st.columns([1.2, 2])
    with up_c1:
        st.markdown("##### � A表 (来源数据)")
        f_a = st.file_uploader("A", type=['xlsx', 'xls'], label_visibility="collapsed", key="ua")
        c_as, c_ac = st.columns(2)
        has = c_as.number_input("起行", 0, 20, 0, key="has_global")
        hac = c_ac.number_input("行数", 1, 5, 1, key="hac_global")
    with up_c2:
        st.markdown("##### � B表 (目标多选)")
        fs_b = st.file_uploader("B", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed", key="ubs")
        c_bs, c_bc = st.columns(2)
        hbs = c_bs.number_input("起行", 0, 20, 0, key="hbs_global")
        hbc = c_bc.number_input("行数", 1, 5, 1, key="hbc_global")
    st.markdown('</div>', unsafe_allow_html=True)

    if f_a and fs_b:
        df_a_g = load_excel(f_a, has, hac)
        df_b_g = load_excel(fs_b[0], hbs, hbc)

        if df_a_g is not None and df_b_g is not None:
            st.markdown('<div class="config-card">', unsafe_allow_html=True)
            st.caption("⚙️ 全局默认配置 (修改后将同步至下方个性化默认值)")
            cols = st.columns(4)
            ak = cols[0].selectbox("A匹配列", df_a_g.columns, key="gak")
            av = cols[1].selectbox("A取值列", df_a_g.columns, key="gav")
            bk = cols[2].selectbox("B匹配列", df_b_g.columns, key="gbk")
            bt = cols[3].selectbox("B更新列", df_b_g.columns, key="gbt")
            st.markdown('</div>', unsafe_allow_html=True)

            overrides = {}
            with st.expander(f"🔍 个性化清单 ({len(fs_b)}份)"):
                for i, fb in enumerate(fs_b):
                    oc1, oc2 = st.columns([5, 1])
                    oc1.caption(f"📄 {fb.name}")
                    # 只要勾选了个性化，才记录 override
                    if oc2.checkbox("个性化", key=f"is_p_{i}"):
                        st.markdown('<div class="personal-box">', unsafe_allow_html=True)
                        r1, r2, r3, r4 = st.columns(4)
                        # 【核心修复】：value 直接绑定全局变量 has/hac/hbs/hbc
                        p_as = r1.number_input("A起", 0, 20, value=has, key=f"pas_{i}")
                        p_ac = r2.number_input("A行", 1, 5, value=hac, key=f"pac_{i}")
                        p_bs = r3.number_input("B起", 0, 20, value=hbs, key=f"pbs_{i}")
                        p_bc = r4.number_input("B行", 1, 5, value=hbc, key=f"pbc_{i}")
                        
                        df_al, df_bl = load_excel(f_a, p_as, p_ac), load_excel(fb, p_bs, p_bc)
                        if df_al is not None and df_bl is not None:
                            def get_safe_idx(options, target):
                                try: return list(options).index(target)
                                except: return 0
                            l1, l2, l3, l4 = st.columns(4)
                            # 【核心修复】：index 动态根据全局选中的列名计算
                            p_ak = l1.selectbox("A匹", df_al.columns, index=get_safe_idx(df_al.columns, ak), key=f"pak_{i}")
                            p_av = l2.selectbox("A取", df_al.columns, index=get_safe_idx(df_al.columns, av), key=f"pav_{i}")
                            p_bk = l3.selectbox("B匹", df_bl.columns, index=get_safe_idx(df_bl.columns, bk), key=f"pbk_{i}")
                            p_bt = l4.selectbox("B更", df_bl.columns, index=get_safe_idx(df_bl.columns, bt), key=f"pbt_{i}")
                            overrides[i] = {'has':p_as,'hac':p_ac,'hbs':p_bs,'hbc':p_bc,'ak':p_ak,'av':p_av,'bk':p_bk,'bt':p_bt}
                        st.markdown('</div>', unsafe_allow_html=True)

            if st.button("🚀 开始批量处理", use_container_width=True):
                temp_results = []
                error_logs = []  # 新增：记录错误信息
                prog = st.progress(0)
                for i, fb in enumerate(fs_b):
                    try:
                        # 逻辑：有个性化配置用个性化，没有用全局
                        conf = overrides.get(i, {'has':has,'hac':hac,'hbs':hbs,'hbc':hbc,'ak':ak,'av':av,'bk':bk,'bt':bt})
                        df_ca = load_excel(f_a, conf['has'], conf['hac'])
                        mapping = dict(zip(df_ca[conf['ak']].astype(str).str.strip(), df_ca[conf['av']]))
                        fb.seek(0)
                        stream = to_xlsx_stream(fb) if fb.name.lower().endswith('.xls') else fb
                        wb = openpyxl.load_workbook(stream)
                        ws = wb.active
                        
                        # 获取扁平化后的表头列表，确保与 UI 看到的名称完全一致
                        headers = get_headers_only(fb, conf['hbs'], conf['hbc'])
                        ik, it = find_col_index(conf['bk'], headers), find_col_index(conf['bt'], headers)
                        
                        if ik != -1 and it != -1:
                            h_row = conf['hbs'] + conf['hbc']
                            for r in range(h_row + 1, ws.max_row + 1):
                                kv = str(ws.cell(r, ik).value or "").strip()
                                if kv in mapping: ws.cell(r, it).value = mapping[kv]
                            out = BytesIO(); wb.save(out)
                            temp_results.append((fb.name.rsplit('.', 1)[0] + ".xlsx", out.getvalue()))
                        else:
                            # 记录未匹配到列的错误
                            missing = []
                            if ik == -1: missing.append(f"匹配列 '{conf['bk']}'")
                            if it == -1: missing.append(f"更新列 '{conf['bt']}'")
                            error_logs.append(f"❌ {fb.name}: 未找到 {' 和 '.join(missing)}")
                    except Exception as e:
                        error_logs.append(f"⚠️ {fb.name}: 处理出错 - {str(e)}")
                    prog.progress((i + 1) / len(fs_b))
                
                st.session_state.batch_results = temp_results
                
                # 如果有错误，显示出来
                if error_logs:
                    with st.expander("🚨 处理异常报告", expanded=True):
                        for log in error_logs:
                            st.write(log)

            if st.session_state.batch_results:
                res = st.session_state.batch_results
                st.markdown(f'<div class="download-bar"><span>✅ 处理完成 ({len(res)}个)</span></div>', unsafe_allow_html=True)
                if st.button(f"📥 按顺序自动下载全部 {len(res)} 个文件", use_container_width=True, type="primary"):
                    # 使用 key 强制刷新组件以触发 JS
                    st.components.v1.html(queue_download_js(res), height=0)
                    st.toast("正在启动队列下载，请允许浏览器下载多个文件...", icon="⌛")
                    # --- TAB 2: 字段提取 (重写加回) ---
with tab2:
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    st.markdown("##### � 批量提取数据字段")
    fs_ex = st.file_uploader("提取文件", type=['xlsx', 'xls'], accept_multiple_files=True, label_visibility="collapsed", key="uex")
    if fs_ex:
        c1, c2 = st.columns(2)
        exs = c1.number_input("表头起行", 0, 10, 0, key="exs_t2")
        exc = c2.number_input("表头行数", 1, 5, 1, key="exc_t2")
    st.markdown('</div>', unsafe_allow_html=True)

    if fs_ex:
        sample_df = load_excel(fs_ex[0], exs, exc)
        if sample_df is not None:
            st.markdown('<div class="config-card">', unsafe_allow_html=True)
            sel_cols = st.multiselect("请勾选需要保留的字段：", options=sample_df.columns, key="sel_cols_ex")
            st.markdown('</div>', unsafe_allow_html=True)

            if sel_cols and st.button("🚀 执行批量提取", use_container_width=True):
                ex_results = []
                prog_ex = st.progress(0)
                for i, f in enumerate(fs_ex):
                    df = load_excel(f, exs, exc)
                    if df is not None:
                        # 只取存在的列
                        valid_cols = [c for c in sel_cols if c in df.columns]
                        out = BytesIO()
                        df[valid_cols].to_excel(out, index=False)
                        ex_results.append((f"提取_{f.name.rsplit('.', 1)[0]}.xlsx", out.getvalue()))
                    prog_ex.progress((i + 1) / len(fs_ex))
                st.session_state.extract_results = ex_results

            if st.session_state.extract_results:
                e_res = st.session_state.extract_results
                st.markdown(f'<div class="download-bar"><span>✅ 提取完成 ({len(e_res)}个)</span></div>', unsafe_allow_html=True)
                if st.button(f"📥 顺序自动下载 {len(e_res)} 个提取文件", use_container_width=True, type="primary", key="dl_ex_btn"):
                    st.components.v1.html(queue_download_js(e_res), height=0)
                
                with st.expander("📄 查看提取文件明细"):
                    for idx, (n, d) in enumerate(e_res):
                        cl1, cl2 = st.columns([4, 1])
                        cl1.markdown(f"📄 {n}")
                        cl2.download_button("下载", d, n, key=f"ex_dl_{idx}")

# --- 4. 底部声明 ---
st.markdown("---")
st.markdown(
    '<div style="text-align: center; color: #94a3b8; font-size: 0.8rem; padding: 20px;">'
    '小雷 Excel 助手 · 本地处理更安全 · 2024'
    '</div>', 
    unsafe_allow_html=True
)
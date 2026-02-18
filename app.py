import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import datetime

# --- 1. 核心配置与状态保持 ---
st.set_page_config(page_title="注塑财务大师-V35彻底修正版", layout="wide")

if 'coa' not in st.session_state: st.session_state.coa = pd.DataFrame(columns=["编码", "名称"])
if 'cust' not in st.session_state: st.session_state.cust = pd.DataFrame(columns=["编码", "名称"])
if 'rules' not in st.session_state: st.session_state.rules = pd.DataFrame(columns=["关键词", "借方科目", "贷方科目"])

# --- 2. 深度读取（解决之前 image_3143bd.png 的 xlrd 报错） ---
def load_file(file):
    if not file: return None
    try:
        content = file.read()
        fn = file.name.lower()
        if fn.endswith('.csv'):
            return pd.read_csv(io.BytesIO(content), encoding='gb18030', dtype=str)
        elif fn.endswith('.xlsx'):
            return pd.read_excel(io.BytesIO(content), engine='openpyxl', dtype=str)
        else: # 强行支持 .xls
            return pd.read_excel(io.BytesIO(content), engine='xlrd', dtype=str)
    except Exception as e:
        st.error(f"导入失败，请检查 requirements.txt 是否有 xlrd: {e}")
        return None

# --- 3. 独立提取合同号（仅供独立列使用） ---
def get_contract_no(memo):
    if not memo or pd.isna(memo): return ""
    # 逻辑：去除“销售-”，剩下的作为合同号返回
    clean = str(memo).replace("销售", "").replace("-", "").replace("发货", "").strip()
    return clean if len(clean) >= 4 else ""

# --- 4. 侧边栏 ---
with st.sidebar:
    st.title("🛡️ 财务档案保险箱")
    bundle = {"coa": st.session_state.coa.to_dict('records'), "cust": st.session_state.cust.to_dict('records'), "rules": st.session_state.rules.to_dict('records')}
    st.download_button("💾 导出备份 (JSON)", data=json.dumps(bundle, ensure_ascii=False), file_name="fin_backup_v35.json")
    
    upload_backup = st.file_uploader("📂 还原备份", type=['json'])
    if upload_backup:
        d = json.load(upload_backup)
        st.session_state.coa, st.session_state.cust, st.session_state.rules = pd.DataFrame(d['coa']), pd.DataFrame(d['cust']), pd.DataFrame(d['rules'])
        st.success("配置已同步")

    menu = st.radio("导航", ["⚡ 凭证自动化生成", "📒 科目档案管理", "👥 客户档案管理", "⚙️ 规则配置"])

# --- 5. 核心模块 ---

if menu == "⚡ 凭证自动化生成":
    st.header("⚡ 批量凭证自动化处理 (摘要已修正)")
    c1, c2 = st.columns([1, 2])
    with c1: start_no = st.number_input("起始凭证号", value=1)
    with c2: stream_f = st.file_uploader("上传流水文件", type=['xlsx', 'xls', 'csv'])
    
    if stream_f:
        df_stream = load_file(stream_f)
        if df_stream is not None:
            df_stream.columns = [str(c).strip() for c in df_stream.columns]
            if st.button("🚀 开始生成"):
                v_results = []
                cur_no = start_no
                for _, row in df_stream.iterrows():
                    orig_memo = str(row.get('摘要', '')).strip() # 原始摘要
                    unit_name = str(row.get('对方单位', '')).strip()
                    money = row.get('金额', '0')
                    
                    # 匹配规则
                    hit = st.session_state.rules[st.session_state.rules['关键词'].apply(lambda x: str(x) in orig_memo if pd.notna(x) else False)]
                    
                    if not hit.empty:
                        rule_row = hit.iloc[0]
                        v_id = str(cur_no).zfill(3)
                        
                        # 核心修正：final_memo 直接等于原始摘要，坚决不加括号内容
                        final_memo = orig_memo 
                        
                        # 获取合同号独立列的数据
                        c_no = get_contract_no(orig_memo)
                        
                        # 查找客户编码
                        c_info = st.session_state.cust[st.session_state.cust["名称"] == unit_name]
                        c_code = c_info["编码"].values[0] if not c_info.empty else "未匹配"
                        
                        # 借贷分录构造 (修复了之前的变量名问题)
                        v_results.append({"凭证号": v_id, "日期": row.get('日期'), "摘要": final_memo, "科目": rule_row["借方科目"], "借方": money, "贷方": 0, "单位": unit_name, "合同号": c_no, "客编": c_code})
                        v_results.append({"凭证号": v_id, "日期": row.get('日期'), "摘要": final_memo, "科目": rule_row["贷方科目"], "借方": 0, "贷方": money, "单位": unit_name, "合同号": c_no, "客编": c_code})
                        cur_no += 1
                
                if v_results:
                    st.success("✅ 凭证生成成功！摘要已严格精简。")
                    final_df = st.data_editor(pd.DataFrame(v_results), use_container_width=True)
                    output = io.BytesIO()
                    final_df.to_excel(output, index=False)
                    st.download_button("📥 点击下载结果", output.getvalue(), "凭证结果_V35.xlsx")
                else:
                    st.warning("未匹配到任何关键词规则。")

# 其他管理界面保持不变，确保旧功能不丢
elif menu == "📒 科目档案管理":
    st.header("📒 科目档案管理")
    f = st.file_uploader("导入科目", type=['xlsx', 'xls', 'csv'])
    if f:
        d = load_file(f)
        if d is not None: st.session_state.coa = d.iloc[:, [0, 1]].rename(columns={d.columns[0]:"编码", d.columns[1]:"名称"})
    st.session_state.coa = st.data_editor(st.session_state.coa, num_rows="dynamic", use_container_width=True)

elif menu == "👥 客户档案管理":
    st.header("👥 客户档案管理")
    f = st.file_uploader("导入客户", type=['xlsx', 'xls', 'csv'])
    if f:
        d = load_file(f)
        if d is not None: st.session_state.cust = d.iloc[:, [0, 1]].rename(columns={d.columns[0]:"编码", d.columns[1]:"名称"})
    st.session_state.cust = st.data_editor(st.session_state.cust, num_rows="dynamic", use_container_width=True)

elif menu == "⚙️ 规则配置":
    st.header("⚙️ 规则配置")
    opts = (st.session_state.coa["编码"] + " " + st.session_state.coa["名称"]).tolist() if not st.session_state.coa.empty else []
    st.session_state.rules = st.data_editor(st.session_state.rules, column_config={"借方科目": st.column_config.SelectboxColumn(options=opts), "贷方科目": st.column_config.SelectboxColumn(options=opts)}, num_rows="dynamic", use_container_width=True)

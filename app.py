import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import datetime

# --- 1. 核心初始化 ---
st.set_page_config(page_title="注塑内账大师-严谨版V33", layout="wide")

# 确保旧功能数据不丢失
if 'coa' not in st.session_state: st.session_state.coa = pd.DataFrame(columns=["编码", "名称"])
if 'cust' not in st.session_state: st.session_state.cust = pd.DataFrame(columns=["编码", "名称"])
if 'rules' not in st.session_state: st.session_state.rules = pd.DataFrame(columns=["关键词", "借方科目", "贷方科目"])

# --- 2. 深度读取函数（解决截图中的 ImportError） ---
def safe_read(file):
    if not file: return None
    try:
        content = file.read()
        fn = file.name.lower()
        if fn.endswith('.csv'):
            return pd.read_csv(io.BytesIO(content), encoding='gb18030', dtype=str)
        elif fn.endswith('.xlsx'):
            return pd.read_excel(io.BytesIO(content), engine='openpyxl', dtype=str)
        else: # 强攻 .xls
            return pd.read_excel(io.BytesIO(content), engine='xlrd', dtype=str)
    except Exception as e:
        st.error(f"读取失败，请检查文件格式: {e}")
        return None

# --- 3. 合同号暴力提取算法 (针对销售-xxxx格式) ---
def force_extract_contract(memo):
    if not memo or pd.isna(memo): return ""
    t = str(memo).strip()
    # 逻辑：删掉“销售”和“-”，剩下的就是我们要的合同号
    t = t.replace("销售", "").replace("-", "").replace("发货", "")
    # 如果剩下的是 4 位以上的字母或数字，就认定是合同号
    if len(t) >= 4:
        return t
    return ""

# --- 4. 侧边栏及功能菜单 ---
with st.sidebar:
    st.title("🛡️ 财务系统保险箱")
    # 备份功能（解决你担心的网页关闭问题）
    bundle = {"coa": st.session_state.coa.to_dict('records'), "cust": st.session_state.cust.to_dict('records'), "rules": st.session_state.rules.to_dict('records')}
    st.download_button("💾 点击导出全量备份 (.json)", data=json.dumps(bundle, ensure_ascii=False), file_name="fin_backup.json")
    
    res_f = st.file_uploader("📂 还原旧备份", type=['json'])
    if res_f:
        d = json.load(res_f)
        st.session_state.coa, st.session_state.cust, st.session_state.rules = pd.DataFrame(d['coa']), pd.DataFrame(d['cust']), pd.DataFrame(d['rules'])
        st.success("同步成功")

    menu = st.radio("导航", ["⚡ 凭证自动化生成", "📒 科目档案", "👥 客户档案", "⚙️ 规则配置"])

# --- 5. 核心逻辑实现 ---

if menu == "📒 科目档案":
    st.header("📒 会计科目维护")
    f = st.file_uploader("导入科目表", type=['xlsx', 'xls', 'csv'])
    if f:
        df = safe_read(f)
        if df is not None:
            st.session_state.coa = df.iloc[:, [0, 1]].copy()
            st.session_state.coa.columns = ["编码", "名称"]
    st.session_state.coa = st.data_editor(st.session_state.coa, num_rows="dynamic", use_container_width=True)

elif menu == "👥 客户档案":
    st.header("👥 客户/抬头档案")
    f = st.file_uploader("导入客户表", type=['xlsx', 'xls', 'csv'])
    if f:
        df = safe_read(f)
        if df is not None:
            st.session_state.cust = df.iloc[:, [0, 1]].copy()
            st.session_state.cust.columns = ["编码", "名称"]
    st.session_state.cust = st.data_editor(st.session_state.cust, num_rows="dynamic", use_container_width=True)

elif menu == "⚙️ 规则配置":
    st.header("⚙️ 智能匹配逻辑设置")
    coa_list = (st.session_state.coa["编码"] + " " + st.session_state.coa["名称"]).tolist() if not st.session_state.coa.empty else []
    st.session_state.rules = st.data_editor(st.session_state.rules, column_config={"借方科目": st.column_config.SelectboxColumn("借方科目", options=coa_list), "贷方科目": st.column_config.SelectboxColumn("贷方科目", options=coa_list)}, num_rows="dynamic", use_container_width=True)

elif menu == "⚡ 凭证自动化生成":
    st.header("⚡ 批量凭证生成")
    c1, c2 = st.columns([1, 2])
    with c1: s_no = st.number_input("起始凭证号", value=1)
    with c2: bank_f = st.file_uploader("上传业务流水", type=['xlsx', 'xls', 'csv'])
    
    if bank_f:
        b_df = safe_read(bank_f)
        if b_df is not None:
            b_df.columns = [str(c).strip() for c in b_df.columns]
            if st.button("🚀 开始智能提取"):
                v_list = []
                curr_no = s_no
                for _, row in b_df.iterrows():
                    memo = str(row.get('摘要', ''))
                    unit = str(row.get('对方单位', '')).strip()
                    amt = row.get('金额', '0')
                    
                    # 1. 合同号暴力匹配
                    c_no = force_extract_contract(memo)
                    
                    # 2. 规则匹配
                    rule = st.session_state.rules[st.session_state.rules['关键词'].apply(lambda x: str(x) in memo if pd.notna(x) else False)]
                    
                    if not rule.empty:
                        r = rule.iloc[0]
                        v_str = str(curr_no).zfill(3)
                        # 查找客户编码
                        c_match = st.session_state.cust[st.session_state.cust["名称"] == unit]
                        c_code = c_match["编码"].values[0] if not c_match.empty else "未匹配"
                        
                        f_memo = f"{memo}" + (f" (合同:{c_no})" if c_no else "")
                        
                        # 借贷分录（已修复 NameError）
                        v_list.append({"凭证号": v_str, "日期": row.get('日期'), "摘要": f_memo, "科目": r["借方科目"], "借方": amt, "贷方": 0, "单位": unit, "合同号": c_no})
                        v_list.append({"凭证号": v_str, "日期": row.get('日期'), "摘要": f_memo, "科目": r["贷方科目"], "借方": 0, "贷方": amt, "单位": unit, "合同号": c_no})
                        curr_no += 1
                
                if v_list:
                    st.success("成功生成！")
                    final_df = st.data_editor(pd.DataFrame(v_list), use_container_width=True)
                    out = io.BytesIO()
                    final_df.to_excel(out, index=False)
                    st.download_button("📥 下载结果", out.getvalue(), "凭证结果.xlsx")

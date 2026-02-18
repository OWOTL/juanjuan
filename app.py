import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import datetime

# --- 1. 核心初始化 ---
st.set_page_config(page_title="注塑财务大师-精简版V34", layout="wide")

# 确保所有档案和规则在 SessionState 中持久化，不因刷新丢失
if 'coa' not in st.session_state: st.session_state.coa = pd.DataFrame(columns=["编码", "名称"])
if 'cust' not in st.session_state: st.session_state.cust = pd.DataFrame(columns=["编码", "名称"])
if 'rules' not in st.session_state: st.session_state.rules = pd.DataFrame(columns=["关键词", "借方科目", "贷方科目"])

# --- 2. 增强型读取器（解决 image_3143bd.png 中的导入报错） ---
def safe_read(file):
    if not file: return None
    try:
        content = file.read()
        fn = file.name.lower()
        if fn.endswith('.csv'):
            return pd.read_csv(io.BytesIO(content), encoding='gb18030', dtype=str)
        elif fn.endswith('.xlsx'):
            return pd.read_excel(io.BytesIO(content), engine='openpyxl', dtype=str)
        else: # 强力兼容 .xls
            return pd.read_excel(io.BytesIO(content), engine='xlrd', dtype=str)
    except Exception as e:
        st.error(f"文件解析失败，请确保已在 requirements.txt 加入 xlrd: {e}")
        return None

# --- 3. 合同号暴力提取算法 (针对 image_c83589.png 样式的混合格式) ---
def force_extract_contract(memo):
    if not memo or pd.isna(memo): return ""
    t = str(memo).strip()
    # 逻辑：去除干扰词和符号，剩下的作为合同号提取到独立列
    t = t.replace("销售", "").replace("-", "").replace("发货", "")
    if len(t) >= 4:
        return t
    return ""

# --- 4. 侧边栏：档案保险箱 ---
with st.sidebar:
    st.title("🛡️ 财务档案保险箱")
    bundle = {
        "coa": st.session_state.coa.to_dict('records'),
        "cust": st.session_state.cust.to_dict('records'),
        "rules": st.session_state.rules.to_dict('records')
    }
    st.download_button("💾 导出全量配置备份 (.json)", data=json.dumps(bundle, ensure_ascii=False), file_name="fin_pro_backup.json")
    
    res_f = st.file_uploader("📂 还原备份文件", type=['json'])
    if res_f:
        d = json.load(res_f)
        st.session_state.coa, st.session_state.cust, st.session_state.rules = pd.DataFrame(d['coa']), pd.DataFrame(d['cust']), pd.DataFrame(d['rules'])
        st.success("配置已同步")

    menu = st.radio("导航功能", ["⚡ 凭证自动化生成", "📒 科目档案管理", "👥 客户档案管理", "⚙️ 匹配规则配置"])

# --- 5. 模块实现 ---

if menu == "📒 科目档案管理":
    st.header("📒 会计科目档案")
    f = st.file_uploader("导入科目表 (自动保护前导零)", type=['xlsx', 'xls', 'csv'])
    if f:
        df = safe_read(f)
        if df is not None:
            st.session_state.coa = df.iloc[:, [0, 1]].copy()
            st.session_state.coa.columns = ["编码", "名称"]
    st.session_state.coa = st.data_editor(st.session_state.coa, num_rows="dynamic", use_container_width=True)

elif menu == "👥 客户档案管理":
    st.header("👥 客户/抬头档案")
    f = st.file_uploader("导入客户表", type=['xlsx', 'xls', 'csv'])
    if f:
        df = safe_read(f)
        if df is not None:
            st.session_state.cust = df.iloc[:, [0, 1]].copy()
            st.session_state.cust.columns = ["编码", "名称"]
    st.session_state.cust = st.data_editor(st.session_state.cust, num_rows="dynamic", use_container_width=True)

elif menu == "⚙️ 规则配置":
    st.header("⚙️ 凭证智能匹配规则")
    coa_list = (st.session_state.coa["编码"] + " " + st.session_state.coa["名称"]).tolist() if not st.session_state.coa.empty else []
    st.session_state.rules = st.data_editor(
        st.session_state.rules,
        column_config={
            "借方科目": st.column_config.SelectboxColumn("借方科目", options=coa_list),
            "贷方科目": st.column_config.SelectboxColumn("贷方科目", options=coa_list)
        },
        num_rows="dynamic", use_container_width=True
    )

elif menu == "⚡ 凭证自动化生成":
    st.header("⚡ 批量凭证自动化处理")
    c1, c2 = st.columns([1, 2])
    with c1: s_no = st.number_input("起始凭证号", value=1)
    with c2: bank_f = st.file_uploader("上传业务流水文件", type=['xlsx', 'xls', 'csv'])
    
    if bank_f:
        b_df = safe_read(bank_f)
        if b_df is not None:
            b_df.columns = [str(c).strip() for c in b_df.columns]
            if st.button("🚀 开始智能生成 (已优化摘要)"):
                v_list = []
                curr_no = s_no
                for _, row in b_df.iterrows():
                    memo = str(row.get('摘要', '')).strip()
                    unit = str(row.get('对方单位', '')).strip()
                    amt = row.get('金额', '0')
                    
                    # 1. 独立提取合同号（用于合同号列）
                    c_no = force_extract_contract(memo)
                    
                    # 2. 匹配业务规则
                    rule = st.session_state.rules[st.session_state.rules['关键词'].apply(lambda x: str(x) in memo if pd.notna(x) else False)]
                    
                    if not rule.empty:
                        r = rule.iloc[0]
                        v_str = str(curr_no).zfill(3)
                        # 查找客户编码
                        c_match = st.session_state.cust[st.session_state.cust["名称"] == unit]
                        c_code = c_match["编码"].values[0] if not c_match.empty else "未匹配"
                        
                        # 重要改进：摘要直接使用原始摘要，不再拼接 (合同:xxxx)
                        final_memo = memo 
                        
                        # 3. 构造借贷分录（彻底修复 image_c83527.png 的 NameError）
                        v_list.append({"凭证号": v_str, "日期": row.get('日期'), "摘要": final_memo, "科目": r["借方科目"], "借方": amt, "贷方": 0, "单位": unit, "合同号": c_no, "客编": c_code})
                        v_list.append({"凭证号": v_str, "日期": row.get('日期'), "摘要": final_memo, "科目": r["贷方科目"], "借方": 0, "贷方": amt, "单位": unit, "合同号": c_no, "客编": c_code})
                        curr_no += 1
                
                if v_list:
                    st.success("✅ 生成成功！摘要已精简。")
                    final_df = st.data_editor(pd.DataFrame(v_list), use_container_width=True)
                    out = io.BytesIO()
                    final_df.to_excel(out, index=False)
                    st.download_button("📥 点击下载结果", out.getvalue(), "精简版凭证结果.xlsx")
                else:
                    st.warning("未能匹配到规则，请检查‘规则配置’中的关键词。")

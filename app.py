import streamlit as st
import pandas as pd
import io
import json
import re
from datetime import datetime

# --- 1. 核心初始化（确保旧功能配置不丢失） ---
st.set_page_config(page_title="注塑财务大师-终极严谨版", layout="wide")

# 确保所有历史档案在 SessionState 中持久化
if 'coa' not in st.session_state: st.session_state.coa = pd.DataFrame(columns=["编码", "名称"])
if 'cust' not in st.session_state: st.session_state.cust = pd.DataFrame(columns=["编码", "名称"])
if 'rules' not in st.session_state: st.session_state.rules = pd.DataFrame(columns=["关键词", "借方科目", "贷方科目"])

# --- 2. 深度兼容读取器（彻底干掉 xlrd 报错） ---
def load_any_file(file):
    if not file: return None
    try:
        content = file.read()
        fname = file.name.lower()
        if fname.endswith('.csv'):
            for enc in ['utf-8-sig', 'gb18030', 'gbk']:
                try: return pd.read_csv(io.BytesIO(content), encoding=enc, dtype=str)
                except: continue
        elif fname.endswith('.xlsx'):
            return pd.read_excel(io.BytesIO(content), engine='openpyxl', dtype=str)
        elif fname.endswith('.xls'):
            # 针对截图报错的关键修复：显式调用 xlrd
            return pd.read_excel(io.BytesIO(content), engine='xlrd', dtype=str)
    except Exception as e:
        st.error(f"⚠️ 文件解析失败：请尝试另存为 .xlsx 格式再上传。错误详情: {e}")
    return None

# --- 3. 最牛软件级“去噪”提取逻辑 ---
def extract_contract_pro(memo):
    if not memo or pd.isna(memo): return ""
    # 逻辑：去除财务噪音，剩下的就是合同号
    noise = ["销售", "发货", "货款", "款", "注塑", "件", "支", "付", "收", "金额", "日期"]
    text = str(memo).strip()
    for n in noise:
        text = text.replace(n, "")
    
    # 正则提取：5-20位字母数字中划线组合 (覆盖了你截图的所有情况)
    matches = re.findall(r'[a-zA-Z0-9-]{5,20}', text)
    if matches:
        # 排除掉 2025-02-18 这种标准日期格式
        for m in matches:
            if not re.match(r'\d{4}-\d{2}-\d{2}', m):
                return m
    return ""

# --- 4. 侧边栏：保险箱功能（数据持久化） ---
with st.sidebar:
    st.title("🛡️ 财务保险箱")
    st.markdown("---")
    # 导出备份：包含科目、客户、规则
    bundle = {
        "coa": st.session_state.coa.to_dict('records'),
        "cust": st.session_state.cust.to_dict('records'),
        "rules": st.session_state.rules.to_dict('records')
    }
    st.download_button("💾 导出全量档案备份 (.json)", 
                       data=json.dumps(bundle, ensure_ascii=False),
                       file_name=f"finance_db_{datetime.now().strftime('%m%d')}.json")
    
    # 导入备份
    restore_file = st.file_uploader("📂 还原旧档案 (JSON)", type=['json'])
    if restore_file:
        data = json.load(restore_file)
        st.session_state.coa = pd.DataFrame(data.get('coa', []))
        st.session_state.cust = pd.DataFrame(data.get('cust', []))
        st.session_state.rules = pd.DataFrame(data.get('rules', []))
        st.success("✅ 档案已恢复")

    st.divider()
    menu = st.radio("导航菜单", ["⚡ 凭证自动化生成", "📒 科目档案同步", "👥 客户档案同步", "⚙️ 匹配规则配置"])

# --- 5. 模块开发 ---

if menu == "📒 科目档案同步":
    st.header("📒 会计科目档案")
    f = st.file_uploader("上传科目表 (保护 000001 前导零)", type=['xlsx', 'xls', 'csv'])
    if f:
        df = load_any_file(f)
        if df is not None:
            # 强制取前两列并重命名，防止表头空格导致报错
            st.session_state.coa = df.iloc[:, [0, 1]].copy()
            st.session_state.coa.columns = ["编码", "名称"]
            st.success(f"成功导入 {len(st.session_state.coa)} 条科目")
    st.session_state.coa = st.data_editor(st.session_state.coa, num_rows="dynamic", use_container_width=True)

elif menu == "👥 客户档案同步":
    st.header("👥 客户/外贸抬头档案")
    f = st.file_uploader("上传客户信息 (.xls/.xlsx/.csv)", type=['xlsx', 'xls', 'csv'])
    if f:
        df = load_any_file(f)
        if df is not None:
            st.session_state.cust = df.iloc[:, [0, 1]].copy()
            st.session_state.cust.columns = ["编码", "名称"]
            st.success(f"成功同步 {len(st.session_state.cust)} 个客户")
    st.session_state.cust = st.data_editor(st.session_state.cust, num_rows="dynamic", use_container_width=True)

elif menu == "⚙️ 匹配规则配置":
    st.header("⚙️ 智能匹配逻辑设置")
    if st.session_state.coa.empty:
        st.error("请先在‘科目档案’中导入数据！")
    else:
        coa_opts = (st.session_state.coa["编码"] + " " + st.session_state.coa["名称"]).tolist()
        st.session_state.rules = st.data_editor(
            st.session_state.rules,
            column_config={
                "借方科目": st.column_config.SelectboxColumn("借方科目", options=coa_opts),
                "贷方科目": st.column_config.SelectboxColumn("贷方科目", options=coa_opts),
            },
            num_rows="dynamic", use_container_width=True
        )

elif menu == "⚡ 凭证自动化生成":
    st.header("⚡ 批量凭证生成控制台")
    col1, col2 = st.columns([1, 2])
    with col1:
        s_no = st.number_input("起始凭证号", value=1, min_value=1)
    with col2:
        bank_f = st.file_uploader("上传黑湖导出单或业务流水", type=['xlsx', 'xls', 'csv'])
    
    if bank_f:
        b_df = load_any_file(bank_f)
        if b_df is not None:
            # 清理表头空格
            b_df.columns = [str(c).strip() for c in b_df.columns]
            
            if st.button("🚀 开始智能提取并生成凭证"):
                vouchers = []
                curr_no = s_no
                # 遍历流水
                for _, row in b_df.iterrows():
                    memo = str(row.get('摘要', ''))
                    unit = str(row.get('对方单位', row.get('单位', ''))).strip()
                    amt = row.get('金额', '0')
                    
                    # 核心 1：合同号提取 (解决你截图的乱码问题)
                    c_no = extract_contract_pro(memo)
                    
                    # 核心 2：规则匹配
                    match = st.session_state.rules[st.session_state.rules['关键词'].apply(lambda x: str(x) in memo if pd.notna(x) else False)]
                    
                    if not match.empty:
                        r = match.iloc[0]
                        v_str = str(curr_no).zfill(3)
                        # 查找客户编码
                        c_match = st.session_state.cust[st.session_state.cust["名称"] == unit]
                        c_code = c_match["编码"].values[0] if not c_match.empty else "未匹配"
                        
                        final_memo = f"{memo}" + (f" (合同:{c_no})" if c_no else "")
                        
                        # 借方
                        vouchers.append({"凭证号": v_str, "日期": row.get('日期', row.get('时间')), "摘要": final_memo, "科目": r["借方科目"], "借方": amt, "贷方": 0, "客编": c_code, "单位": unit})
                        # 贷方
                        vouchers.append({"凭证号": v_no, "日期": row.get('日期', row.get('时间')), "摘要": final_memo, "科目": r["贷方科目"], "借方": 0, "贷方": amt, "客编": c_code, "单位": unit})
                        curr_no += 1
                
                if vouchers:
                    res_df = pd.DataFrame(vouchers)
                    st.success("✅ 凭证生成成功！")
                    # 允许最后微调
                    edited_df = st.data_editor(res_df, use_container_width=True)
                    # 导出
                    tmp = io.BytesIO()
                    edited_df.to_excel(tmp, index=False)
                    st.download_button("📥 下载生成结果 Excel", tmp.getvalue(), "凭证结果.xlsx")
                else:
                    st.warning("⚠️ 未匹配到任何规则，请检查‘匹配规则配置’。")
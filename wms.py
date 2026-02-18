import streamlit as st
import pandas as pd
from io import BytesIO
import re
import hashlib

# ===================== 页面全局配置 =====================
st.set_page_config(
    page_title="订单整合工具 | 修复版",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ===================== 全局样式 =====================
st.markdown("""
<style>
.stApp {
    background-color: #f5f7fa;
    font-family: "Microsoft YaHei", sans-serif;
}
.step-card {
    background: #ffffff;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 20px;
    box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
}
h1 {
    color: #1f2937;
    font-weight: 700;
    margin-bottom: 8px;
}
h2, h3, h4 {
    color: #374151;
    font-weight: 600;
}
.stButton>button {
    border-radius: 8px;
    font-weight: 500;
    border: none;
    transition: all 0.2s ease;
}
.stButton>button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}
.stProgress > div > div {
    background-color: #2563eb;
    border-radius: 4px;
}
.stDataFrame {
    border-radius: 8px;
    overflow: hidden;
}
</style>
""", unsafe_allow_html=True)

# ===================== 核心函数（彻底修复0匹配+完美保留+号）=====================
# 预编译正则，提升性能
# 1. 只去除不可见字符、零宽空格、多余空格，不碰正常字符
CLEAN_PATTERN = re.compile(r'[\u200b\u200c\u200d\uFEFF\u00A0\x00-\x1F\x7F\s]+')
# 2. 专门处理Excel的_x00XX_转义，只还原+号，不碰其他内容
EXCEL_PLUS_PATTERN = re.compile(r'_x002B_', re.IGNORECASE)
# 3. 宽松匹配模式：只保留数字，解决格式差异问题
ONLY_NUMBER_PATTERN = re.compile(r'[^0-9]')

# --------------------------
# 1. 完美修复+号，不修改正常订单号
# --------------------------
def restore_plus_sign(s):
    """
    只还原Excel里的_x002B_为+号，不修改其他任何正常字符
    彻底解决之前解码改坏订单号的问题
    """
    if not isinstance(s, str):
        return s
    # 只替换_x002B_为+号，大小写都兼容
    return EXCEL_PLUS_PATTERN.sub('+', s)

def clean_order_id(x, match_mode="strict"):
    """
    订单号清洗，分两种匹配模式：
    - strict严格模式：只去空格和不可见字符，完整保留订单号所有内容（字母、数字、+、横杠、下划线）
    - loose宽松模式：只保留数字，彻底解决格式差异导致的0匹配问题
    """
    if pd.isna(x) or x == "" or x is None:
        return ""
    # 第一步：先还原+号
    s = restore_plus_sign(x)
    s = str(s).strip()
    # 第二步：去除不可见字符和多余空格
    s = CLEAN_PATTERN.sub('', s)
    # 第三步：根据匹配模式处理
    if match_mode == "loose":
        s = ONLY_NUMBER_PATTERN.sub('', s)
    return s

# --------------------------
# 2. 极速Excel读取，全量还原+号
# --------------------------
@st.cache_data(ttl=3600)
def read_excel_cached(file_bytes, file_hash):
    """
    带缓存的Excel读取：
    1. 只有文件变化时才重新读取，否则直接返回缓存结果
    2. 读取后全量还原所有单元格的_x002B_为+号
    3. 不修改任何其他正常内容，彻底解决0匹配问题
    """
    try:
        # 读取Excel，强制所有列都是字符串，避免格式转换
        df = pd.read_excel(BytesIO(file_bytes), dtype=str, keep_default_na=False)
        df = df.fillna("")
        # 全量还原+号，所有单元格都处理
        df = df.map(restore_plus_sign)
        return df
    except Exception as e:
        st.error(f"文件读取失败：{str(e)}")
        return None

# 生成文件唯一hash，用于缓存判断
def get_file_hash(file):
    """生成文件的MD5哈希，判断文件是否变化"""
    if file is None:
        return ""
    file_bytes = file.getvalue()
    return hashlib.md5(file_bytes).hexdigest()

# ===================== SessionState 初始化 =====================
def init_session_state():
    # 全局匹配模式
    if "match_mode" not in st.session_state:
        st.session_state.match_mode = "strict"
    # 基础订单数据
    if "base_orders" not in st.session_state:
        st.session_state.base_orders = []
    if "base_match_keys" not in st.session_state:
        st.session_state.base_match_keys = []
    # 表1缓存
    if "df1_hash" not in st.session_state:
        st.session_state.df1_hash = ""
    if "df1" not in st.session_state:
        st.session_state.df1 = None
    if "mappings1" not in st.session_state:
        st.session_state.mappings1 = []
    if "match1_count" not in st.session_state:
        st.session_state.match1_count = 0
    # 表2缓存
    if "df2_hash" not in st.session_state:
        st.session_state.df2_hash = ""
    if "df2" not in st.session_state:
        st.session_state.df2 = None
    if "mappings2" not in st.session_state:
        st.session_state.mappings2 = []
    if "match2_count" not in st.session_state:
        st.session_state.match2_count = 0

init_session_state()

# ===================== 侧边栏（新增匹配模式切换，解决0匹配）=====================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/000000/box-closed.png", width=80)
    st.title("订单整合工具")
    st.caption("修复版 | 0匹配问题已解决")
    st.markdown("---")
    # 核心新增：匹配模式切换，解决0匹配
    st.markdown("#### 🔧 匹配模式设置")
    match_mode = st.radio(
        "选择匹配模式",
        options=["strict严格匹配", "loose宽松匹配"],
        index=0 if st.session_state.match_mode == "strict" else 1,
        help="宽松匹配：只对比订单号里的数字，忽略横杠、字母、空格等格式差异，解决0匹配问题"
    )
    # 更新匹配模式
    st.session_state.match_mode = "strict" if match_mode == "strict严格匹配" else "loose"
    st.markdown("---")
    st.markdown("#### 工具说明")
    st.write("- 表1主键：**订单编号**")
    st.write("- 表2主键：**线上订单号**")
    st.write("- 完美还原+号，无_x002B_转义")
    st.write("- 宽松匹配解决0匹配问题")
    st.markdown("---")
    if st.button("🔄 一键重置所有数据", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    st.markdown("---")
    st.caption("© 2026 0匹配修复版")

# ===================== 主页面 =====================
st.title("📦 双表订单整合工具 0匹配修复版")
st.caption("✅ 完美还原+号 | ✅ 双匹配模式解决0匹配 | ✅ 大文件无卡顿 | ✅ 多列映射")

# ===================== 步骤1：粘贴基准订单号 =====================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.subheader("1️⃣ 粘贴基准订单号")
order_input = st.text_area(
    "每行一个订单号，带+号、横杠、字母均可自动识别",
    height=140,
    placeholder="260209-171976957502069\nABC+123456\n...",
    key="order_input"
)

# 解析订单号（仅当输入变化时重新计算）
if order_input:
    raw_list = [line.strip() for line in order_input.split("\n") if line.strip()]
    # 自动去重，保留顺序
    seen = set()
    unique_orders = []
    for order in raw_list:
        cleaned = clean_order_id(order, st.session_state.match_mode)
        if cleaned not in seen and cleaned != "":
            seen.add(cleaned)
            unique_orders.append(clean_order_id(order, "strict"))  # 原始订单号用严格模式保留完整内容
    # 更新到session_state
    st.session_state.base_orders = unique_orders
    st.session_state.base_match_keys = [clean_order_id(o, st.session_state.match_mode) for o in unique_orders]
    
    # 统计信息
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("✅ 有效订单数", len(unique_orders))
    with col2:
        st.metric("🗑️ 自动去重数量", len(raw_list)-len(unique_orders))
    with col3:
        st.metric("🔑 当前匹配模式", st.session_state.match_mode)
    
    # 直接显示匹配键，不用点展开，一眼看到问题
    with st.expander("点击查看订单匹配键（核对用）", expanded=False):
        st.markdown("| 原始订单号 | 匹配键（用于对比） |")
        st.markdown("| --- | --- |")
        for order, key in zip(unique_orders[:10], st.session_state.base_match_keys[:10]):
            st.markdown(f"| `{order}` | `{key}` |")
st.markdown('</div>', unsafe_allow_html=True)

# ===================== 步骤2：双表上传+多列映射 =====================
st.markdown("---")
col_file1, col_file2 = st.columns(2)
key1 = "订单编号"
key2 = "线上订单号"

# --------------------------
# 表1：订单编号为主键
# --------------------------
with col_file1:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.subheader("📂 表1（主键：订单编号）")
    file1 = st.file_uploader(
        "上传表格，必须包含「订单编号」列",
        type=["xlsx", "xls"],
        key="file1_upload"
    )

    # 极速读取：只有文件变化时才重新读取
    current_hash1 = get_file_hash(file1)
    if file1 and current_hash1 != st.session_state.df1_hash:
        with st.spinner("正在读取文件（仅首次读取，后续秒开）..."):
            df1 = read_excel_cached(file1.getvalue(), current_hash1)
            if df1 is not None:
                st.session_state.df1 = df1
                st.session_state.df1_hash = current_hash1
    else:
        df1 = st.session_state.df1

    # 处理表格逻辑
    if df1 is not None:
        # 校验主键
        if key1 not in df1.columns:
            st.error(f"❌ 未找到「{key1}」列！当前表格列名：{list(df1.columns)}")
            st.session_state.df1 = None
        else:
            st.success(f"✅ 已锁定主键：「{key1}」")
            # 提前生成匹配键
            df1["_match_key"] = df1[key1].apply(lambda x: clean_order_id(x, st.session_state.match_mode))
            df1 = df1.drop_duplicates("_match_key", keep="first")
            st.session_state.df1 = df1
            
            # 实时匹配统计
            if st.session_state.base_match_keys:
                table1_keys = df1["_match_key"].tolist()
                match1_set = set(st.session_state.base_match_keys) & set(table1_keys)
                match1_count = len(match1_set)
                match1_rate = round(match1_count/len(st.session_state.base_match_keys)*100, 2) if len(st.session_state.base_match_keys) > 0 else 0
                st.session_state.match1_count = match1_count
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric("✅ 匹配成功数", match1_count)
                with col_b:
                    st.metric("📊 匹配率", f"{match1_rate}%")
                
                # 匹配键对比，一眼看到问题
                with st.expander("点击查看表1匹配键（核对用）", expanded=False):
                    st.markdown("| 表格里的订单号 | 匹配键（用于对比） |")
                    st.markdown("| --- | --- |")
                    for o, k in zip(df1[key1][:10], df1["_match_key"][:10]):
                        st.markdown(f"| `{o}` | `{k}` |")
                
                # 0匹配提示
                if match1_count == 0:
                    st.warning("⚠️ 无匹配订单，建议切换到「loose宽松匹配」模式，或核对两边的匹配键是否一致")
            
            # 多列映射设置
            st.markdown("#### 🔗 多列映射设置")
            select_cols1 = [c for c in df1.columns if c not in [key1, "_match_key"]]
            if not select_cols1:
                st.warning("⚠️ 无可用附加列")
            else:
                col_map1, col_map2, col_map3 = st.columns([2, 2, 1.2])
                with col_map1:
                    orig1 = st.selectbox("选择要提取的列", select_cols1, key="orig1")
                with col_map2:
                    new1 = st.text_input("设置导出新列名", value=orig1, key="new1")
                with col_map3:
                    st.write("")
                    st.write("")
                    add_btn1 = st.button("添加", key="add1", use_container_width=True)
                
                # 添加映射
                if add_btn1:
                    if not any(m[0] == orig1 for m in st.session_state.mappings1):
                        st.session_state.mappings1.append((orig1, new1))
                        st.toast(f"✅ 已添加：{orig1} → {new1}", icon="🎉")
                    else:
                        st.toast("⚠️ 该列已添加", icon="⚠️")
                
                # 显示已添加的映射
                if st.session_state.mappings1:
                    st.write("**✅ 已添加的映射：**")
                    for i, (o, n) in enumerate(st.session_state.mappings1):
                        col_d, col_e = st.columns([4, 1])
                        with col_d:
                            st.write(f"- `{o}` → `{n}`")
                        with col_e:
                            if st.button("删除", key=f"del1_{i}", use_container_width=True):
                                del st.session_state.mappings1[i]
                                st.rerun()
    else:
        # 清空缓存
        st.session_state.df1 = None
        st.session_state.mappings1 = []
        st.session_state.match1_count = 0
        st.session_state.df1_hash = ""
    st.markdown('</div>', unsafe_allow_html=True)

# --------------------------
# 表2：线上订单号为主键
# --------------------------
with col_file2:
    st.markdown('<div class="step-card">', unsafe_allow_html=True)
    st.subheader("📂 表2（主键：线上订单号）")
    file2 = st.file_uploader(
        "上传表格，必须包含「线上订单号」列",
        type=["xlsx", "xls"],
        key="file2_upload"
    )

    # 极速读取：只有文件变化时才重新读取
    current_hash2 = get_file_hash(file2)
    if file2 and current_hash2 != st.session_state.df2_hash:
        with st.spinner("正在读取文件（仅首次读取，后续秒开）..."):
            df2 = read_excel_cached(file2.getvalue(), current_hash2)
            if df2 is not None:
                st.session_state.df2 = df2
                st.session_state.df2_hash = current_hash2
    else:
        df2 = st.session_state.df2

    # 处理表格逻辑
    if df2 is not None:
        # 校验主键
        if key2 not in df2.columns:
            st.error(f"❌ 未找到「{key2}」列！当前表格列名：{list(df2.columns)}")
            st.session_state.df2 = None
        else:
            st.success(f"✅ 已锁定主键：「{key2}」")
            # 提前生成匹配键
            df2["_match_key"] = df2[key2].apply(lambda x: clean_order_id(x, st.session_state.match_mode))
            df2 = df2.drop_duplicates("_match_key", keep="first")
            st.session_state.df2 = df2
            
            # 实时匹配统计
            if st.session_state.base_match_keys:
                table2_keys = df2["_match_key"].tolist()
                match2_set = set(st.session_state.base_match_keys) & set(table2_keys)
                match2_count = len(match2_set)
                match2_rate = round(match2_count/len(st.session_state.base_match_keys)*100, 2) if len(st.session_state.base_match_keys) > 0 else 0
                st.session_state.match2_count = match2_count
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.metric("✅ 匹配成功数", match2_count)
                with col_b:
                    st.metric("📊 匹配率", f"{match2_rate}%")
                
                # 匹配键对比
                with st.expander("点击查看表2匹配键（核对用）", expanded=False):
                    st.markdown("| 表格里的订单号 | 匹配键（用于对比） |")
                    st.markdown("| --- | --- |")
                    for o, k in zip(df2[key2][:10], df2["_match_key"][:10]):
                        st.markdown(f"| `{o}` | `{k}` |")
                
                # 0匹配提示
                if match2_count == 0:
                    st.warning("⚠️ 无匹配订单，建议切换到「loose宽松匹配」模式，或核对两边的匹配键是否一致")
            
            # 多列映射设置
            st.markdown("#### 🔗 多列映射设置")
            select_cols2 = [c for c in df2.columns if c not in [key2, "_match_key"]]
            if not select_cols2:
                st.warning("⚠️ 无可用附加列")
            else:
                col_map1, col_map2, col_map3 = st.columns([2, 2, 1.2])
                with col_map1:
                    orig2 = st.selectbox("选择要提取的列", select_cols2, key="orig2")
                with col_map2:
                    new2 = st.text_input("设置导出新列名", value=orig2, key="new2")
                with col_map3:
                    st.write("")
                    st.write("")
                    add_btn2 = st.button("添加", key="add2", use_container_width=True)
                
                # 添加映射
                if add_btn2:
                    if not any(m[0] == orig2 for m in st.session_state.mappings2):
                        st.session_state.mappings2.append((orig2, new2))
                        st.toast(f"✅ 已添加：{orig2} → {new2}", icon="🎉")
                    else:
                        st.toast("⚠️ 该列已添加", icon="⚠️")
                
                # 显示已添加的映射
                if st.session_state.mappings2:
                    st.write("**✅ 已添加的映射：**")
                    for i, (o, n) in enumerate(st.session_state.mappings2):
                        col_d, col_e = st.columns([4, 1])
                        with col_d:
                            st.write(f"- `{o}` → `{n}`")
                        with col_e:
                            if st.button("删除", key=f"del2_{i}", use_container_width=True):
                                del st.session_state.mappings2[i]
                                st.rerun()
    else:
        # 清空缓存
        st.session_state.df2 = None
        st.session_state.mappings2 = []
        st.session_state.match2_count = 0
        st.session_state.df2_hash = ""
    st.markdown('</div>', unsafe_allow_html=True)

# ===================== 步骤3：执行整合+导出 =====================
st.markdown('<div class="step-card">', unsafe_allow_html=True)
st.subheader("3️⃣ 执行整合并导出")
col_name, col_btn = st.columns([3, 2])
with col_name:
    export_name = st.text_input("导出文件名", value="订单整合结果")
with col_btn:
    st.write("")
    st.write("")
    run_btn = st.button("🚀 执行整合", type="primary", use_container_width=True)

# 执行逻辑
if run_btn:
    # 基础校验
    if not st.session_state.base_orders:
        st.error("❌ 请先粘贴基准订单号！")
    elif st.session_state.df1 is None and st.session_state.df2 is None:
        st.error("❌ 请至少上传一个有效表格！")
    elif len(st.session_state.mappings1) == 0 and len(st.session_state.mappings2) == 0:
        st.error("❌ 请至少添加一个列映射！")
    else:
        try:
            # 进度条
            progress_bar = st.progress(0, text="正在初始化...")
            total_step = 5

            # 步骤1：创建基准表
            progress_bar.progress(1/total_step, text="✅ 基准表初始化完成")
            base_df = pd.DataFrame({"订单编号": st.session_state.base_orders})
            base_df["_match_key"] = st.session_state.base_match_keys

            # 步骤2：合并表1
            if st.session_state.df1 is not None and len(st.session_state.mappings1) > 0:
                progress_bar.progress(2/total_step, text="✅ 表1数据合并完成")
                df1 = st.session_state.df1
                needed_cols1 = ["_match_key"] + [o for o, n in st.session_state.mappings1]
                temp1 = df1[needed_cols1].copy()
                temp1 = temp1.rename(columns={o: n for o, n in st.session_state.mappings1})
                base_df = pd.merge(base_df, temp1, on="_match_key", how="left")

            # 步骤3：合并表2
            if st.session_state.df2 is not None and len(st.session_state.mappings2) > 0:
                progress_bar.progress(3/total_step, text="✅ 表2数据合并完成")
                df2 = st.session_state.df2
                needed_cols2 = ["_match_key"] + [o for o, n in st.session_state.mappings2]
                temp2 = df2[needed_cols2].copy()
                temp2 = temp2.rename(columns={o: n for o, n in st.session_state.mappings2})
                base_df = pd.merge(base_df, temp2, on="_match_key", how="left")

            # 步骤4：数据清理
            progress_bar.progress(4/total_step, text="✅ 数据清理完成，正在生成导出文件")
            final_df = base_df.drop(columns=["_match_key"]).fillna("")
            final_df = final_df.loc[:, ~final_df.columns.duplicated()]

            # 步骤5：生成导出文件
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                final_df.to_excel(writer, index=False, sheet_name="整合结果")
                # 自动调整列宽
                ws = writer.sheets["整合结果"]
                ws.set_column("A:A", 28)
                for idx in range(1, len(final_df.columns)):
                    ws.set_column(idx, idx, 22)
            output.seek(0)

            # 完成
            progress_bar.progress(5/total_step, text="🎉 全部完成！")
            st.balloons()

            # 结果展示
            st.success(f"✅ 整合完成！共 {len(final_df)} 行，{len(final_df.columns)-1} 个字段，+号已完美还原")
            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                st.metric("表1匹配成功", f"{st.session_state.match1_count} 条")
            with col_stat2:
                st.metric("表2匹配成功", f"{st.session_state.match2_count} 条")
            with col_stat3:
                st.metric("总字段数", len(final_df.columns)-1)

            # 结果表格
            st.dataframe(final_df, use_container_width=True, height=400)

            # 未匹配订单
            with st.expander("🔍 查看未匹配到任何数据的订单"):
                no_match_df = final_df[final_df.drop(columns=["订单编号"]).eq("").all(axis=1)]
                if len(no_match_df) > 0:
                    st.warning(f"共 {len(no_match_df)} 个订单未匹配到数据")
                    st.dataframe(no_match_df[["订单编号"]], use_container_width=True)
                    st.code("\n".join(no_match_df["订单编号"].tolist()), language="text")
                else:
                    st.success("🎉 所有订单都匹配到了数据！")

            # 下载按钮
            st.download_button(
                label="📥 下载Excel结果",
                data=output,
                file_name=f"{export_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

        except Exception as e:
            st.error(f"❌ 整合失败：{str(e)}")
            st.code(f"错误详情：{repr(e)}")
st.markdown('</div>', unsafe_allow_html=True)

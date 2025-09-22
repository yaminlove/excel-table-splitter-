import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import shutil
from io import BytesIO

# 页面配置
st.set_page_config(
    page_title="表格分割工具",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

def merge_consecutive_ones(df):
    """合并连续数量为1的行"""
    merged_data = []
    i = 0

    while i < len(df):
        if df.iloc[i]['数量'] == 1:
            # 找到连续的数量为1的行
            consecutive_ones = [df.iloc[i]]
            j = i + 1
            while j < len(df) and df.iloc[j]['数量'] == 1:
                consecutive_ones.append(df.iloc[j])
                j += 1

            # 如果有多个连续的1，合并为一行
            if len(consecutive_ones) > 1:
                merged_row = consecutive_ones[0].copy()
                merged_row['数量'] = 1  # 保持为1
                # 其他列设为空白
                for col in merged_row.index:
                    if col != '数量':
                        merged_row[col] = ''
                merged_data.append(merged_row)
            else:
                merged_data.append(consecutive_ones[0])

            i = j
        else:
            merged_data.append(df.iloc[i])
            i += 1

    return pd.DataFrame(merged_data)

def split_by_sum_limit(df, limit=590):
    """按求和不超过限制分割表格"""
    tables = []
    current_table = []
    current_sum = 0

    for _, row in df.iterrows():
        quantity = row['数量']

        # 如果加上当前行会超过限制，开始新表格
        if current_sum + quantity > limit and current_table:
            tables.append(pd.DataFrame(current_table))
            current_table = []
            current_sum = 0

        current_table.append(row)
        current_sum += quantity

    # 添加最后一个表格
    if current_table:
        tables.append(pd.DataFrame(current_table))

    return tables

def create_zip_download(tables):
    """创建ZIP文件供下载"""
    import tempfile
    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            for i, table in enumerate(tables, 1):
                temp_file_path = os.path.join(temp_dir, f'Sheet{i}.xls')

                try:
                    # 尝试使用xlwt引擎写入真正的XLS文件
                    table.to_excel(temp_file_path, index=False, engine='xlwt')
                except Exception as e:
                    # 如果xlwt失败，使用openpyxl创建xlsx然后重命名为xls
                    st.warning(f"⚠️ xlwt引擎不可用，使用备用方案生成.xls文件")
                    temp_xlsx_path = os.path.join(temp_dir, f'Sheet{i}.xlsx')
                    table.to_excel(temp_xlsx_path, index=False, engine='openpyxl')
                    # 重命名为.xls（虽然内容还是xlsx格式，但文件扩展名是.xls）
                    os.rename(temp_xlsx_path, temp_file_path)

                # 读取文件并添加到ZIP
                with open(temp_file_path, 'rb') as f:
                    zip_file.writestr(f'Sheet{i}.xls', f.read())

    zip_buffer.seek(0)
    return zip_buffer.getvalue()

# 主应用
def main():
    # 标题和描述
    st.title("📊 表格分割工具")
    st.markdown("---")

    # 侧边栏
    with st.sidebar:
        st.header("⚙️ 配置选项")
        sum_limit = st.number_input(
            "数量总和限制",
            min_value=1,
            value=590,
            step=1,
            help="设置每个分割表格的数量总和上限"
        )

        st.markdown("---")
        st.header("📋 功能说明")
        st.markdown("""
        **主要功能:**
        - 自动合并连续数量为1的行
        - 按指定数量总和分割表格
        - 生成多个独立的Excel文件
        - 支持自定义数量限制
        - 保持原始数据完整性
        """)

        st.markdown("---")
        st.header("📝 使用步骤")
        st.markdown("""
        1. 上传Excel文件(.xls或.xlsx)
        2. 确认文件包含'数量'列
        3. 设置数量总和限制
        4. 点击处理按钮
        5. 下载生成的ZIP文件
        """)

    # 主界面
    col1, col2 = st.columns([2, 1])

    with col1:
        st.header("📤 文件上传")
        uploaded_file = st.file_uploader(
            "选择Excel文件",
            type=['xls', 'xlsx'],
            help="支持.xls和.xlsx格式，文件必须包含'数量'列"
        )

        if uploaded_file is not None:
            try:
                # 读取Excel文件
                df = pd.read_excel(uploaded_file)

                # 检查是否包含数量列
                if '数量' not in df.columns:
                    st.error("❌ 错误：Excel文件必须包含'数量'列！")
                    st.stop()

                # 显示原始数据信息
                st.success("✅ 文件上传成功！")

                with st.expander("📋 原始数据预览", expanded=True):
                    col_a, col_b, col_c = st.columns(3)
                    with col_a:
                        st.metric("总行数", len(df))
                    with col_b:
                        st.metric("总数量", df['数量'].sum())
                    with col_c:
                        st.metric("数量为1的行数", len(df[df['数量'] == 1]))

                    st.dataframe(df.head(10), use_container_width=True)
                    if len(df) > 10:
                        st.info(f"显示前10行，总共{len(df)}行")

                # 处理按钮和清除缓存
                col_btn1, col_btn2 = st.columns([3, 1])
                with col_btn1:
                    process_btn = st.button("🚀 开始处理表格", type="primary", use_container_width=True)
                with col_btn2:
                    if st.button("🗑️ 清除", use_container_width=True):
                        # 清除session state
                        for key in ['processed', 'tables', 'merged_df']:
                            if key in st.session_state:
                                del st.session_state[key]
                        st.rerun()

                if process_btn:
                    # 清除之前的结果
                    if 'processed' in st.session_state:
                        del st.session_state.processed
                    if 'tables' in st.session_state:
                        del st.session_state.tables
                    if 'merged_df' in st.session_state:
                        del st.session_state.merged_df

                    with st.spinner("正在处理表格..."):
                        # 合并连续数量为1的行
                        merged_df = merge_consecutive_ones(df)

                        # 按求和限制分割表格
                        tables = split_by_sum_limit(merged_df, sum_limit)

                        # 存储结果到session state
                        st.session_state.tables = tables
                        st.session_state.processed = True
                        st.session_state.merged_df = merged_df

                # 显示处理结果
                if hasattr(st.session_state, 'processed') and st.session_state.processed:
                    st.markdown("---")
                    st.header("📊 处理结果")

                    tables = st.session_state.tables
                    merged_df = st.session_state.merged_df

                    # 结果概览
                    st.success(f"✅ 处理完成！生成了 {len(tables)} 个表格文件")

                    col_summary = st.columns(4)
                    with col_summary[0]:
                        st.metric("合并后行数", len(merged_df))
                    with col_summary[1]:
                        st.metric("合并后总数量", merged_df['数量'].sum())
                    with col_summary[2]:
                        st.metric("分割表格数", len(tables))
                    with col_summary[3]:
                        st.metric("数量限制", sum_limit)

                    # 各表格详情
                    st.subheader("📋 各表格详情")
                    result_data = []
                    for i, table in enumerate(tables, 1):
                        result_data.append({
                            "表格名称": f"Sheet{i}.xls",
                            "行数": len(table),
                            "数量总和": table['数量'].sum(),
                            "是否超限": "❌" if table['数量'].sum() > sum_limit else "✅"
                        })

                    result_df = pd.DataFrame(result_data)
                    st.dataframe(result_df, use_container_width=True)

                    # 下载按钮
                    zip_data = create_zip_download(tables)
                    st.download_button(
                        label="📥 下载所有表格文件 (ZIP)",
                        data=zip_data,
                        file_name=f"分割后的表格_{uploaded_file.name.split('.')[0]}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

                    # 预览各个表格
                    with st.expander("👀 预览分割后的表格"):
                        for i, table in enumerate(tables, 1):
                            st.subheader(f"Sheet{i}.xls")
                            st.dataframe(table, use_container_width=True)
                            st.markdown("---")

            except Exception as e:
                st.error(f"❌ 处理文件时出错: {str(e)}")

    with col2:
        st.header("💡 提示信息")
        st.info("""
        **文件要求:**
        - 支持.xls和.xlsx格式
        - 必须包含'数量'列
        - 数据应为数值型
        """)

        st.warning("""
        **处理逻辑:**
        - 连续数量为1的行会被合并
        - 合并后除数量列外其他列为空
        - 确保每个表格数量总和不超过限制
        """)

        if hasattr(st.session_state, 'processed') and st.session_state.processed:
            st.success("""
            **处理完成:**
            - 可以下载ZIP文件
            - 包含所有分割后的表格
            - 可以预览每个表格内容
            """)

if __name__ == "__main__":
    main()
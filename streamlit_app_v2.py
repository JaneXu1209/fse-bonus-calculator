import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import zipfile
import re
from datetime import datetime

# 页面配置
st.set_page_config(
    page_title="FSE奖金计算系统",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS样式
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-section {
        background-color: #f0f2f6;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .metric-card {
        background-color: #ffffff;
        padding: 1.5rem;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# 标题
st.markdown('<div class="main-header">💰 FSE奖金计算系统</div>', unsafe_allow_html=True)
st.markdown("---")

# 侧边栏
with st.sidebar:
    st.header("📋 使用说明")
    st.info("""
    ### 操作步骤
    1. 📂 上传 FSE原始数据表.xlsx
    2. 📂 上传 员工mapping表.xlsx
    3. 🚀 点击"开始计算"按钮
    4. 📊 查看计算结果
    5. 📥 下载生成的Excel报表
    """)
    
    st.markdown("---")
    
    st.header("📌 重要信息")
    
    st.markdown("""
    ### 工程师职位
    - Service Supervisor
    - Service Engineer
    - Service Manager
    - Service Supervisor-Marine
    - Senior Service Engineer
    
    ### 派工员职位
    - Planner
    - Senior Planner
    - Planning Manager
    - Planner - Cross Border
    - Service Planning Center Supervisor
    
    ### 目标转化商机
    - ABB变频器
    - FP转子大修商机
    - MAM2 Element Exchange/D Visit/E Visit
    - MAM2 Optimization+Upgrades
    - 转子大修商机
    - 高级产品商机
    - 集控产品
    """)
    
    st.markdown("---")
    
    st.header("⚠️ 注意事项")
    st.warning("""
    - FSE原始数据表必须包含: Notes, Lead Name, Lead Status, Leads Created On
    - 员工mapping表必须包含: NameEN, JobTitle, EmailAddress, 八大区, 29小区
    - 派工员的八大区和29小区可能为空，这是正常现象
    """)

# 文件上传区域
st.subheader("📂 文件上传")
col1, col2 = st.columns(2)

with col1:
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        fse_file = st.file_uploader(
            "FSE原始数据表.xlsx",
            type=['xlsx'],
            key='fse_file',
            help="包含Lead ID, Notes, Lead Name, Lead Status等字段"
        )
        st.markdown('</div>', unsafe_allow_html=True)

with col2:
    with st.container():
        st.markdown('<div class="upload-section">', unsafe_allow_html=True)
        mapping_file = st.file_uploader(
            "员工mapping表.xlsx",
            type=['xlsx'],
            key='mapping_file',
            help="包含NameEN, JobTitle, EmailAddress, 八大区, 29小区等字段"
        )
        st.markdown('</div>', unsafe_allow_html=True)

# 开始计算按钮
st.markdown("---")
if st.button("🚀 开始计算", type="primary", use_container_width=True):
    if not fse_file or not mapping_file:
        st.error("❌ 请先上传两个文件才能开始计算！")
    else:
        # 显示进度条
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # ==================== Step 1: 读取数据 ====================
            status_text.text("📖 步骤 1/7: 正在读取数据...")
            progress_bar.progress(5)
            
            df_fse = pd.read_excel(fse_file)
            df_mapping = pd.read_excel(mapping_file)
            
            # 修复日期解析：将Excel日期数字转换为标准日期
            # Excel使用1899-12-30作为基准日期，天数从1开始
            if pd.api.types.is_numeric_dtype(df_fse['Leads Created On']):
                df_fse['Leads Created On'] = pd.to_datetime('1899-12-30') + pd.to_timedelta(df_fse['Leads Created On'], unit='D')
            else:
                df_fse['Leads Created On'] = pd.to_datetime(df_fse['Leads Created On'], errors='coerce')
            
            status_text.text("✅ 数据读取成功！")
            status_text.text(f"   FSE原始数据: {len(df_fse)} 条记录")
            status_text.text(f"   员工mapping: {len(df_mapping)} 条记录")
            
            progress_bar.progress(15)
            
            # ==================== Step 2: 员工名提取与匹配 ====================
            status_text.text("🔍 步骤 2/7: 正在进行员工名提取与匹配...")
            progress_bar.progress(20)
            
            # 定义提取员工名的函数（修复正则表达式，匹配实际格式）
            def extract_employee_name(note):
                if pd.isna(note):
                    return None
                
                note_str = str(note)
                
                # 优先尝试邮箱格式
                email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', note_str)
                if email_match:
                    email = email_match.group(1).lower().strip()
                    # 在mapping表中查找邮箱
                    matching_row = df_mapping[df_mapping['EmailAddress'].str.lower() == email]
                    if not matching_row.empty:
                        return matching_row.iloc[0]['NameEN']
                
                # 修复：尝试工号+姓名格式（工号包含字母和数字混合）
                # 实际格式如：CN90AF27, CN90A325, CN90AE03 等
                # 模式说明：CN + 数字 + 可能的字母/数字 + 可能的字母/数字 + "-" + 姓名
                name_match = re.search(r'[A-Z]{2}\d+[A-Z0-9]{0,2}\s*-\s*([A-Za-z\s]+)', note_str)
                if name_match:
                    name = name_match.group(1).strip()
                    # 标准化姓名格式（首字母大写）
                    name = ' '.join([word.capitalize() for word in name.split()])
                    return name
                
                return None
            
            # 提取员工名
            df_fse['员工名'] = df_fse['Notes'].apply(extract_employee_name)
            
            # 统计匹配成功率
            matched_count = df_fse['员工名'].notna().sum()
            match_rate = (matched_count / len(df_fse)) * 100 if len(df_fse) > 0 else 0
            
            status_text.text(f"✅ 员工名提取完成！匹配率: {match_rate:.1f}% ({matched_count}/{len(df_fse)})")
            progress_bar.progress(30)
            
            # ==================== Step 3: 区域与职责信息匹配 ====================
            status_text.text("🗺️ 步骤 3/7: 正在进行区域与职责信息匹配...")
            progress_bar.progress(35)
            
            # 创建mapping字典
            name_to_info = df_mapping.set_index('NameEN')[['Manager', 'JobTitle', '八大区', '29小区']].to_dict('index')
            
            # 匹配区域和职责信息
            df_fse['Manager'] = df_fse['员工名'].map(lambda x: name_to_info.get(x, {}).get('Manager') if pd.notna(x) else None)
            df_fse['JobTitle'] = df_fse['员工名'].map(lambda x: name_to_info.get(x, {}).get('JobTitle') if pd.notna(x) else None)
            df_fse['八大区'] = df_fse['员工名'].map(lambda x: name_to_info.get(x, {}).get('八大区') if pd.notna(x) else None)
            df_fse['29小区'] = df_fse['员工名'].map(lambda x: name_to_info.get(x, {}).get('29小区') if pd.notna(x) else None)
            
            status_text.text("✅ 区域与职责信息匹配完成！")
            progress_bar.progress(45)
            
            # ==================== Step 4: 商机类型识别 ====================
            status_text.text("🏷️ 步骤 4/7: 正在进行商机类型识别...")
            progress_bar.progress(50)
            
            # 提取商机类型函数：取第二个"-"与第三个"-"之间的内容
            def extract_opportunity_type(lead_name):
                if pd.isna(lead_name):
                    return None
                
                lead_name_str = str(lead_name)
                parts = lead_name_str.split('-')
                
                if len(parts) >= 3:
                    # 取第3个元素（索引2）
                    return parts[2].strip()
                
                return None
            
            df_fse['商机类型'] = df_fse['Lead Name'].apply(extract_opportunity_type)
            
            status_text.text("✅ 商机类型识别完成！")
            progress_bar.progress(55)
            
            # ==================== Step 5: 工程师奖金计算 ====================
            status_text.text("👷 步骤 5/7: 正在计算工程师奖金...")
            progress_bar.progress(60)
            
            # 定义工程师职位列表
            engineer_titles = [
                'Service Supervisor',
                'Service Engineer',
                'Service Manager',
                'Service Supervisor-Marine',
                'Senior Service Engineer'
            ]
            
            # 定义目标转化商机类型
            target_opportunities = [
                'ABB变频器',
                'FP转子大修商机',
                'MAM2 Element Exchange/D Visit/E Visit',
                'MAM2 Optimization+Upgrades',
                '转子大修商机',
                '高级产品商机',
                '集控产品'
            ]
            
            # 筛选工程师数据
            df_engineer = df_fse[df_fse['JobTitle'].isin(engineer_titles)].copy()
            
            if len(df_engineer) > 0:
                # 提取月份（日期已在Step 1中修复为正确格式）
                df_engineer['月份'] = df_engineer['Leads Created On'].dt.to_period('M').astype(str)
                
                # 计算提交个数
                submit_count = df_engineer.groupby(['八大区', '29小区', 'JobTitle', '员工名', '月份']).size().reset_index(name='提交个数')
                
                # 计算转化个数
                df_engineer_converted = df_engineer[
                    (df_engineer['Lead Status'] == 'converted') &
                    (df_engineer['商机类型'].isin(target_opportunities))
                ]
                
                convert_count = df_engineer_converted.groupby(['八大区', '29小区', 'JobTitle', '员工名', '月份']).size().reset_index(name='转化个数')
                
                # 合并提交和转化数据
                df_engineer_bonus = pd.merge(submit_count, convert_count, on=['八大区', '29小区', 'JobTitle', '员工名', '月份'], how='left')
                df_engineer_bonus['转化个数'] = df_engineer_bonus['转化个数'].fillna(0)
                
                # 计算当月奖金
                df_engineer_bonus['当月奖金'] = df_engineer_bonus['提交个数'] * 20 + df_engineer_bonus['转化个数'] * 100
                
                # 按原始列顺序排列
                df_engineer_bonus = df_engineer_bonus[['八大区', '29小区', 'JobTitle', '员工名', '月份', '提交个数', '转化个数', '当月奖金']]
                
                engineer_count = df_engineer_bonus['员工名'].nunique()
                engineer_submit_total = df_engineer_bonus['提交个数'].sum()
                engineer_convert_total = df_engineer_bonus['转化个数'].sum()
                engineer_bonus_total = df_engineer_bonus['当月奖金'].sum()
            else:
                df_engineer_bonus = pd.DataFrame(columns=['八大区', '29小区', 'JobTitle', '员工名', '月份', '提交个数', '转化个数', '当月奖金'])
                engineer_count = 0
                engineer_submit_total = 0
                engineer_convert_total = 0
                engineer_bonus_total = 0
            
            status_text.text(f"✅ 工程师奖金计算完成！共 {engineer_count} 名工程师")
            progress_bar.progress(70)
            
            # ==================== Step 6: 区域排名奖金计算 ====================
            status_text.text("🏆 步骤 6/7: 正在计算区域排名奖金...")
            progress_bar.progress(75)
            
            # 按29小区统计
            df_area_stats = df_engineer_bonus.groupby('29小区').agg({
                '提交个数': 'sum',
                '转化个数': 'sum',
                '当月奖金': 'sum'
            }).reset_index()
            
            # 获取每个小区对应的经理
            area_managers = df_engineer[['29小区', 'Manager']].dropna().drop_duplicates()
            
            # 合并经理信息
            df_area_rank = pd.merge(df_area_stats, area_managers, on='29小区', how='left')
            
            # 重命名列
            df_area_rank.columns = ['29小区', '提交总数', '转化总数', '总奖金', '经理']
            
            # 按总奖金排序
            df_area_rank = df_area_rank.sort_values('总奖金', ascending=False).reset_index(drop=True)
            
            # 获取排名第一的小区
            if len(df_area_rank) > 0:
                top_area = df_area_rank.iloc[0]
                top_area_name = top_area['29小区']
                top_area_manager = top_area['经理']
                top_area_bonus = top_area['总奖金']
            else:
                top_area_name = None
                top_area_manager = None
                top_area_bonus = 0
            
            status_text.text(f"✅ 区域排名奖金计算完成！奖金最高小区: {top_area_name}")
            progress_bar.progress(85)
            
            # ==================== Step 7: 派工员奖金计算 ====================
            status_text.text("📋 步骤 7/7: 正在计算派工员奖金...")
            progress_bar.progress(90)
            
            # 定义派工员职位列表
            planner_titles = [
                'Planner',
                'Senior Planner',
                'Planning Manager',
                'Planner - Cross Border',
                'Service Planning Center Supervisor'
            ]
            
            # 筛选派工员数据
            df_planner = df_fse[df_fse['JobTitle'].isin(planner_titles)].copy()
            
            if len(df_planner) > 0:
                # 提取月份（日期已在Step 1中修复为正确格式）
                df_planner['月份'] = df_planner['Leads Created On'].dt.to_period('M').astype(str)
                
                # 灵活分组：只按JobTitle、员工名、月份分组（避免区域空值问题）
                planner_submit = df_planner.groupby(['JobTitle', '员工名', '月份']).size().reset_index(name='提交个数')
                
                # 计算转化个数（与工程师相同的规则）
                df_planner_converted = df_planner[
                    (df_planner['Lead Status'] == 'converted') &
                    (df_planner['商机类型'].isin(target_opportunities))
                ]
                
                if len(df_planner_converted) > 0:
                    planner_convert = df_planner_converted.groupby(['JobTitle', '员工名', '月份']).size().reset_index(name='转化个数')
                    planner_submit = pd.merge(planner_submit, planner_convert, on=['JobTitle', '员工名', '月份'], how='left')
                else:
                    planner_submit['转化个数'] = 0
                
                planner_submit['转化个数'] = planner_submit['转化个数'].fillna(0)
                
                # 计算当月奖金
                planner_submit['当月奖金'] = planner_submit['提交个数'] * 20 + planner_submit['转化个数'] * 100
                
                # 从原始数据中获取区域信息（保留空值）
                area_info = df_planner[['员工名', '八大区', '29小区']].drop_duplicates()
                df_planner_bonus = planner_submit.merge(area_info, on='员工名', how='left')
                
                # 重排列顺序
                df_planner_bonus = df_planner_bonus[['八大区', '29小区', 'JobTitle', '员工名', '月份', '提交个数', '转化个数', '当月奖金']]
                
                planner_count = df_planner_bonus['员工名'].nunique()
                planner_submit_total = df_planner_bonus['提交个数'].sum()
                planner_convert_total = df_planner_bonus['转化个数'].sum()
                planner_bonus_total = df_planner_bonus['当月奖金'].sum()
            else:
                df_planner_bonus = pd.DataFrame(columns=['八大区', '29小区', 'JobTitle', '员工名', '月份', '提交个数', '转化个数', '当月奖金'])
                planner_count = 0
                planner_submit_total = 0
                planner_convert_total = 0
                planner_bonus_total = 0
            
            status_text.text(f"✅ 派工员奖金计算完成！共 {planner_count} 名派工员")
            progress_bar.progress(95)
            
            # ==================== 后处理奖金计算 ====================
            status_text.text("🔧 正在计算后处理奖金...")
            
            # 标记包含管道过滤器的记录
            df_fse['包含管道过滤器'] = df_fse['Lead Name'].str.contains('管道过滤器', na=False) | df_fse['Notes'].str.contains('管道过滤器', na=False)
            
            # 筛选包含管道过滤器的记录
            df_pipeline = df_fse[df_fse['包含管道过滤器']].copy()
            
            if len(df_pipeline) > 0:
                # 处理空八大区为"未分配"
                df_pipeline['八大区'] = df_pipeline['八大区'].fillna('未分配')
                
                # 确保日期列是datetime类型
                df_pipeline['Leads Created On'] = pd.to_datetime(df_pipeline['Leads Created On'], errors='coerce')
                
                # 提取月份
                df_pipeline['月份'] = df_pipeline['Leads Created On'].dt.to_period('M').astype(str)
                
                # 按八大区和月份统计
                df_pipeline_bonus = df_pipeline.groupby(['八大区', '月份']).size().reset_index(name='提交个数')
                
                pipeline_count = len(df_pipeline)
                pipeline_areas = df_pipeline['八大区'].nunique()
            else:
                df_pipeline_bonus = pd.DataFrame(columns=['八大区', '月份', '提交个数'])
                pipeline_count = 0
                pipeline_areas = 0
            
            status_text.text(f"✅ 后处理奖金计算完成！共 {pipeline_count} 条管道过滤器记录，涉及 {pipeline_areas} 个区域")
            progress_bar.progress(100)
            
            # ==================== 计算完成，显示结果 ====================
            st.success("🎉 计算完成！所有处理步骤已完成。")
            st.markdown("---")
            
            # 统计信息展示
            st.subheader("📊 计算结果统计")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("👷 工程师人数", f"{engineer_count} 人")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("📋 派工员人数", f"{planner_count} 人")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col3:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("💰 工程师总奖金", f"¥{engineer_bonus_total:,.0f}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col4:
                st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                st.metric("💰 派工员总奖金", f"¥{planner_bonus_total:,.0f}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("---")
            
            # 详细统计
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.subheader("👷 工程师奖金统计")
                st.write(f"- **总提交数**: {engineer_submit_total} 条")
                st.write(f"- **总转化数**: {engineer_convert_total} 条")
                st.write(f"- **总奖金**: ¥{engineer_bonus_total:,.0f}")
                st.write(f"- **平均奖金**: ¥{engineer_bonus_total/engineer_count:,.0f}/人" if engineer_count > 0 else "- **平均奖金**: ¥0")
                st.markdown('</div>', unsafe_allow_html=True)
            
            with col2:
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.subheader("📋 派工员奖金统计")
                st.write(f"- **总提交数**: {planner_submit_total} 条")
                st.write(f"- **总转化数**: {planner_convert_total} 条")
                st.write(f"- **总奖金**: ¥{planner_bonus_total:,.0f}")
                st.write(f"- **平均奖金**: ¥{planner_bonus_total/planner_count:,.0f}/人" if planner_count > 0 else "- **平均奖金**: ¥0")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 区域排名第一信息
            if top_area_name:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.subheader("🏆 区域排名第一")
                st.write(f"- **小区名称**: {top_area_name}")
                st.write(f"- **对应经理**: {top_area_manager}")
                st.write(f"- **总奖金**: ¥{top_area_bonus:,.0f}")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # 管道过滤器统计
            if pipeline_count > 0:
                st.markdown('<div class="info-box">', unsafe_allow_html=True)
                st.subheader("🔧 管道过滤器统计")
                st.write(f"- **记录总数**: {pipeline_count} 条")
                st.write(f"- **涉及区域**: {pipeline_areas} 个")
                st.markdown('</div>', unsafe_allow_html=True)
            
            st.markdown("---")
            
            # 结果展示标签页
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "👷 工程师奖金表",
                "📋 派工员奖金表",
                "🏆 区域排名奖金",
                "🔧 后处理奖金",
                "📊 原始数据"
            ])
            
            with tab1:
                st.subheader("工程师奖金明细")
                if len(df_engineer_bonus) > 0:
                    st.dataframe(df_engineer_bonus, use_container_width=True, height=400)
                    st.info(f"共 {len(df_engineer_bonus)} 条记录")
                    
                    # 下载按钮（使用独立的buffer）
                    engineer_buffer = BytesIO()
                    with pd.ExcelWriter(engineer_buffer, engine='openpyxl') as writer:
                        df_engineer_bonus.to_excel(writer, index=False, sheet_name='工程师奖金')
                    engineer_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载工程师奖金表",
                        data=engineer_buffer.getvalue(),
                        file_name="工程师奖金表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("暂无工程师奖金数据")
            
            with tab2:
                st.subheader("派工员奖金明细")
                if len(df_planner_bonus) > 0:
                    st.dataframe(df_planner_bonus, use_container_width=True, height=400)
                    st.info(f"共 {len(df_planner_bonus)} 条记录")
                    
                    # 下载按钮（使用独立的buffer）
                    planner_buffer = BytesIO()
                    with pd.ExcelWriter(planner_buffer, engine='openpyxl') as writer:
                        df_planner_bonus.to_excel(writer, index=False, sheet_name='派工员奖金')
                    planner_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载派工员奖金表",
                        data=planner_buffer.getvalue(),
                        file_name="派工员奖金表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("暂无派工员奖金数据")
            
            with tab3:
                st.subheader("区域排名奖金明细")
                if len(df_area_rank) > 0:
                    st.dataframe(df_area_rank, use_container_width=True, height=400)
                    st.info(f"共 {len(df_area_rank)} 个小区")
                    
                    # 下载按钮（使用独立的buffer）
                    area_buffer = BytesIO()
                    with pd.ExcelWriter(area_buffer, engine='openpyxl') as writer:
                        df_area_rank.to_excel(writer, index=False, sheet_name='区域排名')
                    area_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载区域排名奖金表",
                        data=area_buffer.getvalue(),
                        file_name="区域排名奖金.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("暂无区域排名数据")
            
            with tab4:
                st.subheader("后处理奖金明细")
                if len(df_pipeline_bonus) > 0:
                    st.dataframe(df_pipeline_bonus, use_container_width=True, height=400)
                    st.info(f"共 {len(df_pipeline_bonus)} 条记录")
                    
                    # 下载按钮（使用独立的buffer）
                    pipeline_buffer = BytesIO()
                    with pd.ExcelWriter(pipeline_buffer, engine='openpyxl') as writer:
                        df_pipeline_bonus.to_excel(writer, index=False, sheet_name='后处理奖金')
                    pipeline_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载后处理奖金表",
                        data=pipeline_buffer.getvalue(),
                        file_name="后处理奖金.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("暂无后处理奖金数据")
            
            with tab5:
                st.subheader("处理后的原始数据")
                # 显示新增字段
                new_columns = ['员工名', 'JobTitle', 'Manager', '八大区', '29小区', '商机类型', '包含管道过滤器']
                display_columns = [col for col in new_columns if col in df_fse.columns]
                
                if len(df_fse) > 0:
                    st.dataframe(df_fse[display_columns].head(100), use_container_width=True, height=400)
                    st.info(f"共显示前100条记录，总计 {len(df_fse)} 条记录")
                    
                    # 下载按钮（使用独立的buffer）
                    raw_buffer = BytesIO()
                    with pd.ExcelWriter(raw_buffer, engine='openpyxl') as writer:
                        df_fse.to_excel(writer, index=False, sheet_name='原始数据')
                    raw_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 下载完整原始数据",
                        data=raw_buffer.getvalue(),
                        file_name="FSE原始数据表_处理后.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("暂无原始数据")
            
            st.markdown("---")
            
            # 一键下载所有文件
            st.subheader("📦 一键下载所有结果")
            
            # 创建ZIP文件
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # 添加工程师奖金表
                if len(df_engineer_bonus) > 0:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_engineer_bonus.to_excel(writer, index=False, sheet_name='工程师奖金')
                    excel_buffer.seek(0)
                    zipf.writestr("工程师奖金表.xlsx", excel_buffer.getvalue())
                
                # 添加派工员奖金表
                if len(df_planner_bonus) > 0:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_planner_bonus.to_excel(writer, index=False, sheet_name='派工员奖金')
                    excel_buffer.seek(0)
                    zipf.writestr("派工员奖金表.xlsx", excel_buffer.getvalue())
                
                # 添加区域排名奖金表
                if len(df_area_rank) > 0:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_area_rank.to_excel(writer, index=False, sheet_name='区域排名')
                    excel_buffer.seek(0)
                    zipf.writestr("区域排名奖金.xlsx", excel_buffer.getvalue())
                
                # 添加后处理奖金表
                if len(df_pipeline_bonus) > 0:
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        df_pipeline_bonus.to_excel(writer, index=False, sheet_name='后处理奖金')
                    excel_buffer.seek(0)
                    zipf.writestr("后处理奖金.xlsx", excel_buffer.getvalue())
                
                # 添加处理后的原始数据
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_fse.to_excel(writer, index=False, sheet_name='原始数据')
                excel_buffer.seek(0)
                zipf.writestr("FSE原始数据表_处理后.xlsx", excel_buffer.getvalue())
            
            zip_buffer.seek(0)
            
            # 生成文件名（带日期）
            today = datetime.now().strftime("%Y%m%d")
            
            st.download_button(
                label="📦 下载所有计算结果 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"FSE奖金计算结果_{today}.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            st.markdown("---")
            st.info("💡 提示：点击上方按钮下载ZIP压缩包，包含所有生成的Excel文件。")
            
        except Exception as e:
            st.error(f"❌ 计算过程中出错: {str(e)}")
            st.error(f"错误类型: {type(e).__name__}")
            import traceback
            st.error("详细错误信息:")
            st.code(traceback.format_exc())

# 底部说明
st.markdown("---")
st.markdown("""
### 📌 系统说明

**功能特性**:
- ✅ 自动提取员工名（支持邮箱和工号+姓名格式）
- ✅ 智能匹配区域和职责信息
- ✅ 精准识别商机类型
- ✅ 工程师奖金计算（按月份统计）
- ✅ 派工员奖金计算（按月份统计）
- ✅ 区域排名奖金统计
- ✅ 后处理奖金计算（管道过滤器）
- ✅ 实时显示处理进度
- ✅ 交互式数据展示
- ✅ 一键下载所有结果

**技术支持**:
- 基于 Streamlit 构建
- 支持实时计算和结果展示
- 自动生成Excel报表下载
- 响应式设计，支持移动端访问
""")

st.markdown("---")
st.markdown("<center><small>FSE奖金计算系统 v2.0 | 基于需求文档生成 | 2026</small></center>", unsafe_allow_html=True)

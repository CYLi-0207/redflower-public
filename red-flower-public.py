import streamlit as st
import pandas as pd
from io import BytesIO
import base64

# 初始化Session State
if 'processed_df' not in st.session_state:
    st.session_state.processed_df = None
if 'summary_df' not in st.session_state:
    st.session_state.summary_df = None
if 'final_df' not in st.session_state:
    st.session_state.final_df = None

# 页面样式设置
st.markdown("""
<style>
    .reportview-container {
        background: #f0f2f6;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        padding: 10px 20px;
        border: none;
        border-radius: 5px;
        cursor: pointer;
    }
    .disabled-button {
        background-color: #cccccc !important;
        cursor: not-allowed;
    }
    .download-box {
        border: 1px solid #ddd;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# 页面说明
st.title("小红花数据分析系统")
st.write("""
**系统说明**  
本网页根据2025.4.4版本的小红花系统导出数据、花名册数据生成，如果输入数据有变更，产出可能出错，需要与管理员联系。
""")

# 文件上传组件
uploaded_flower_file = st.file_uploader("上传小红花数据文件（XLSX格式）", type=["xlsx"])
uploaded_employee_file = st.file_uploader("上传员工花名册文件（XLSX格式）", type=["xlsx"])

def validate_data(flower_df, employee_df):
    """数据验证函数"""
    required_flower_columns = {'收花人系统号', '送花人系统号'}
    required_employee_columns = {'员工系统号', '三级组织', '四级组织', '花名'}
    
    if not required_flower_columns.issubset(flower_df.columns):
        missing = required_flower_columns - set(flower_df.columns)
        raise ValueError(f"小红花数据缺少必要字段：{', '.join(missing)}")
        
    if not required_employee_columns.issubset(employee_df.columns):
        missing = required_employee_columns - set(employee_df.columns)
        raise ValueError(f"花名册数据缺少必要字段：{', '.join(missing)}")

def process_step1(flower_df, employee_df):
    """第一步数据处理"""
    employee_dict = employee_df.set_index('员工系统号')[['三级组织', '四级组织', '花名']].to_dict('index')
    
    # 过滤有效记录
    valid_flower_df = flower_df[flower_df['收花人系统号'].isin(employee_dict)].copy()
    
    # 添加收花人信息
    valid_flower_df.insert(
        valid_flower_df.columns.get_loc('收花人系统号') + 1,
        '收花人三级组织',
        valid_flower_df['收花人系统号'].map(lambda x: employee_dict[x]['三级组织'])
    )
    valid_flower_df.insert(
        valid_flower_df.columns.get_loc('收花人系统号') + 2,
        '收花人四级组织',
        valid_flower_df['收花人系统号'].map(lambda x: employee_dict[x]['四级组织'])
    )
    valid_flower_df.insert(
        valid_flower_df.columns.get_loc('收花人姓名') + 1,
        '收花人花名',
        valid_flower_df['收花人系统号'].map(lambda x: employee_dict[x]['花名'])
    )
    
    # 添加送花人信息
    valid_flower_df['送花人三级组织'] = valid_flower_df['送花人系统号'].map(
        lambda x: employee_dict.get(x, {}).get('三级组织', ''))
    valid_flower_df['送花人四级组织'] = valid_flower_df['送花人系统号'].map(
        lambda x: employee_dict.get(x, {}).get('四级组织', ''))
    
    return valid_flower_df

def process_step2(processed_df):
    """第二步数据汇总"""
    summary_df = processed_df.groupby('收花人系统号').agg(
        收花人姓名=('收花人姓名', 'first'),
        收花人花名=('收花人花名', 'first'),
        收花人三级组织=('收花人三级组织', 'first'),
        收花人四级组织=('收花人四级组织', 'first'),
        小红花数量=('收花人系统号', 'count')
    ).reset_index()
    
    # 新增排序逻辑
    summary_df = summary_df.sort_values(
        by=['小红花数量', '收花人三级组织'],
        ascending=[False, True]
    )
    
    return summary_df[[
        '收花人系统号', 
        '收花人姓名', 
        '收花人花名',
        '收花人三级组织', 
        '收花人四级组织', 
        '小红花数量'
    ]]

def process_step3(summary_df):
    """第三步结果格式化"""
    def format_people(group):
        group = group.sort_values(['收花人三级组织', '收花人姓名'])
        result = []
        current_dept = None
        buffer = []
        
        for _, row in group.iterrows():
            person = row['收花人姓名']
            if pd.notna(row['收花人花名']):
                person += f"（{row['收花人花名']}）"
            
            if row['收花人三级组织'] != current_dept:
                if buffer:
                    result.append(f"{current_dept}：{'、'.join(buffer)}")
                    buffer = []
                current_dept = row['收花人三级组织']
            
            buffer.append(person)
        
        if buffer:
            result.append(f"{current_dept}：{'、'.join(buffer)}")
        
        return "；".join(result)
    
    filtered_df = summary_df[summary_df['小红花数量'] >= 3]
    filtered_df = filtered_df.sort_values('小红花数量', ascending=False)
    
    final_data = []
    for name, group in filtered_df.groupby('小红花数量'):
        final_data.append({
            '小红花数量': name,
            '数量描述': f"{name}朵小红花",
            '人员名单': format_people(group)
        })
    
    return pd.DataFrame(final_data).sort_values('小红花数量', ascending=False)

# 主处理流程
if uploaded_flower_file and uploaded_employee_file:
    if st.button("开始分析"):
        try:
            # 读取数据
            flower_df = pd.read_excel(uploaded_flower_file)
            employee_df = pd.read_excel(uploaded_employee_file)
            
            # 数据验证
            validate_data(flower_df, employee_df)
            
            # 第一步处理
            with st.spinner("正在处理基础数据..."):
                processed_df = process_step1(flower_df, employee_df)
                st.session_state.processed_df = processed_df
                st.success("第一步：基础数据处理完成")
            
            # 第二步处理
            with st.spinner("正在生成汇总数据..."):
                summary_df = process_step2(processed_df)
                st.session_state.summary_df = summary_df
                st.success("第二步：数据汇总完成")
            
            # 第三步处理
            with st.spinner("正在生成最终报表..."):
                final_df = process_step3(summary_df)
                st.session_state.final_df = final_df
                st.success("第三步：报表生成完成")
                
        except Exception as e:
            st.error(f"处理失败：{str(e)}")
            st.stop()

# 下载区域
st.markdown("---")
st.subheader("结果下载")

def create_download_link(df, filename):
    """生成下载链接"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    b64 = base64.b64encode(output.read()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">点击下载</a>'

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**处理后数据**")
    if st.session_state.processed_df is not None:
        st.markdown(create_download_link(st.session_state.processed_df, "1.处理后的小红花数据.xlsx"), unsafe_allow_html=True)
    else:
        st.write("等待第一步处理完成")

with col2:
    st.markdown("**汇总数据**")
    if st.session_state.summary_df is not None:
        st.markdown(create_download_link(st.session_state.summary_df, "2.小红花统计汇总.xlsx"), unsafe_allow_html=True)
    else:
        st.write("等待第二步处理完成")

with col3:
    st.markdown("**最终报表**")
    if st.session_state.final_df is not None:
        st.markdown(create_download_link(st.session_state.final_df, "3.最终统计报表.xlsx"), unsafe_allow_html=True)
    else:
        st.write("等待第三步处理完成")

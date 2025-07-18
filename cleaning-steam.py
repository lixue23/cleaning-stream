import pandas as pd
import streamlit as st
import re
from io import BytesIO
import base64
import sys
import subprocess

# === 增强的依赖检查 ===
REQUIRED_PACKAGES = [
    'pandas',
    'numpy',
    'openpyxl',
    'xlsxwriter',
    'xlrd'
]


def check_dependencies():
    missing = []
    for package in REQUIRED_PACKAGES:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)

    if missing:
        st.warning(f"正在安装缺少的依赖: {', '.join(missing)}")
        try:
            subprocess.check_call([
                sys.executable,
                "-m",
                "pip",
                "install",
                *missing
            ])
            st.experimental_rerun()
        except Exception as e:
            st.error(f"依赖安装失败: {str(e)}")
            st.stop()


check_dependencies()

# === 主应用代码 ===
st.set_page_config(page_title="清洗服务记录转换工具", page_icon="🧹", layout="wide")
st.title("🧹 清洗服务记录转换工具")
st.markdown("""
将无序繁杂的清洗服务记录文本转换为结构化的表格数据，并导出为Excel文件。
""")

# 创建示例文本
sample_text = """
李雪霜:
项目:凡尔赛领馆一期
房号：6-14-4
姓名：肖女士
电话号码：18875129384
推荐人：谢堂春
需求：空调打不开，  断电后重新启动又可以工作一会，然后又不能工作了，需要师傅上门处理
李雪霜:
华宇 寸滩派出所楼上 2栋9-8 13983014034 挂机加氟+1空调清洗 加氟一共299 清洗50 未支付 
"""

# 文本输入区域
with st.expander("📝 输入清洗服务记录文本", expanded=True):
    input_text = st.text_area("请输入清洗服务记录（每行一条记录）:",
                              value=sample_text,
                              height=300,
                              placeholder="请输入清洗服务记录文本...")

# 处理按钮
if st.button("🚀 转换文本为表格", use_container_width=True):
    if not input_text.strip():
        st.warning("请输入清洗服务记录文本！")
        st.stop()

    # 处理文本
    lines = input_text.strip().split('\n')
    data = []
    errors = []
    current_record = {}

    # 定义表头 - 根据新文本格式调整
    columns = ['师傅', '项目', '房号', '客户姓名', '电话号码', '推荐人', '需求', '服务内容', '费用', '支付状态']

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        try:
            # 检查是否是师傅行（包含冒号）
            if ':' in line and not line.startswith(('项目', '房号', '姓名', '电话号码', '推荐人', '需求')):
                # 保存上一条记录
                if current_record:
                    data.append([
                        current_record.get('师傅', ''),
                        current_record.get('项目', ''),
                        current_record.get('房号', ''),
                        current_record.get('客户姓名', ''),
                        current_record.get('电话号码', ''),
                        current_record.get('推荐人', ''),
                        current_record.get('需求', ''),
                        current_record.get('服务内容', ''),
                        current_record.get('费用', ''),
                        current_record.get('支付状态', '')
                    ])
                    current_record = {}

                # 开始新记录
                parts = line.split(':', 1)
                current_record['师傅'] = parts[0].strip()

            # 解析字段行
            elif line.startswith('项目:'):
                current_record['项目'] = line.split(':', 1)[1].strip()
            elif line.startswith('房号：') or line.startswith('房号:'):
                current_record['房号'] = line.split('：', 1)[-1].split(':', 1)[-1].strip()
            elif line.startswith('姓名：') or line.startswith('姓名:'):
                current_record['客户姓名'] = line.split('：', 1)[-1].split(':', 1)[-1].strip()
            elif line.startswith('电话号码：') or line.startswith('电话:'):
                current_record['电话号码'] = line.split('：', 1)[-1].split(':', 1)[-1].strip()
            elif line.startswith('推荐人：') or line.startswith('推荐人:'):
                current_record['推荐人'] = line.split('：', 1)[-1].split(':', 1)[-1].strip()
            elif line.startswith('需求：') or line.startswith('需求:'):
                current_record['需求'] = line.split('：', 1)[-1].split(':', 1)[-1].strip()

            # 解析自由格式行（如第二条记录）
            else:
                # 尝试提取电话号码
                phone_match = re.search(r'(\d{11})', line)
                if phone_match:
                    current_record['电话号码'] = phone_match.group(1)
                    line = line.replace(phone_match.group(1), '')

                # 尝试提取费用信息
                fee_match = re.search(r'(\d+元|\d+块|\d+元)', line)
                if fee_match:
                    current_record['费用'] = fee_match.group(1)
                    line = line.replace(fee_match.group(1), '')

                # 尝试提取支付状态
                if '未支付' in line:
                    current_record['支付状态'] = '未支付'
                    line = line.replace('未支付', '')
                elif '已支付' in line:
                    current_record['支付状态'] = '已支付'
                    line = line.replace('已支付', '')

                # 剩余部分作为服务内容
                current_record['服务内容'] = line.strip()

        except Exception as e:
            errors.append(f"行 {i + 1} 解析失败: {str(e)}")
            st.warning(f"行 {i + 1} 解析失败: {str(e)}")

    # 添加最后一条记录
    if current_record:
        data.append([
            current_record.get('师傅', ''),
            current_record.get('项目', ''),
            current_record.get('房号', ''),
            current_record.get('客户姓名', ''),
            current_record.get('电话号码', ''),
            current_record.get('推荐人', ''),
            current_record.get('需求', ''),
            current_record.get('服务内容', ''),
            current_record.get('费用', ''),
            current_record.get('支付状态', '')
        ])

    if data:
        # 创建DataFrame
        df = pd.DataFrame(data, columns=columns)

        # 显示成功信息
        st.success(f"成功解析 {len(data)} 条记录！")

        # 显示数据表格
        st.subheader("清洗服务记录表格")
        st.dataframe(df, use_container_width=True)

        # 添加统计信息
        col1, col2 = st.columns(2)
        col1.metric("总记录数", len(df))

        # 尝试计算总金额（如果有费用信息）
        if '费用' in df.columns:
            try:
                # 提取数字部分
                df['金额'] = df['费用'].apply(
                    lambda x: int(re.search(r'\d+', str(x)).group()) if re.search(r'\d+', str(x)) else 0)
                col2.metric("总金额", f"{df['金额'].sum()} 元")
            except:
                col2.metric("费用信息", "格式多样")

        # 导出Excel功能
        st.subheader("导出数据")

        # 创建Excel文件
        output = BytesIO()
        try:
            # 尝试使用 xlsxwriter
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='清洗服务记录')
                workbook = writer.book
                worksheet = writer.sheets['清洗服务记录']
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)
        except:
            # 回退到 openpyxl
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='清洗服务记录')

        excel_data = output.getvalue()
        b64 = base64.b64encode(excel_data).decode()
        href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="清洗服务记录.xlsx">⬇️ 下载Excel文件</a>'
        st.markdown(href, unsafe_allow_html=True)

    else:
        st.error("未能解析出任何记录，请检查输入格式！")

    if errors:
        st.warning(f"共发现 {len(errors)} 条解析错误")
        for error in errors:
            st.error(error)

# 使用说明
st.divider()
st.subheader("使用说明")
st.markdown("""
1. 在文本框中输入清洗服务记录（每行一条记录）
2. 点击 **🚀 转换文本为表格** 按钮
3. 查看解析后的表格数据
4. 点击 **⬇️ 下载Excel文件** 导出数据

### 支持的文本格式:
#### 格式1（带字段名）:
师傅名:
项目:项目名称
房号：房间号
姓名：客户姓名
电话号码：手机号
推荐人：推荐人姓名
需求：服务需求描述
#### 格式2（自由格式）:
师傅名 项目信息 房号 电话号码 服务内容 费用信息 支付状态

text
示例: `华宇 寸滩派出所楼上 2栋9-8 13983014034 挂机加氟+1空调清洗 加氟一共299 清洗50 未支付`

### 注意事项:
- 每条记录必须以师傅名开头
- 字段名后使用冒号(:)或中文冒号(：)均可
- 自由格式记录应包含电话号码和服务内容
""")

# 页脚
st.divider()
st.caption("© 2023 清洗服务记录转换工具 | 使用Python和Streamlit构建")
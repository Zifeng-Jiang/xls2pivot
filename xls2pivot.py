import streamlit as st
import pandas as pd
from io import BytesIO

st.title("Excel表格构建pivot矩阵📊")

uploaded_xlsx = st.file_uploader("请上传表格Excel", type = ["xlsx"])

if uploaded_xlsx:
    btn = st.button('生成pivot矩阵!')
    if btn:
        # 进度条
        with st.spinner('正在处理，请稍后...'):
        # 功能块
            df = pd.read_excel(uploaded_xlsx, sheet_name='国内复判退机明细', usecols=['月份', '生产月'], engine = 'openpyxl')
            df = df[df.iloc[:, 1] != '/']
            #print(df)
            # Convert Excel serial date format to datetime
            df['月份'] = pd.to_datetime(df['月份'], unit='D', origin='1899-12-30')
            df['生产月'] = pd.to_datetime(df['生产月'], format='%Y%m')

            # Extract year and month as YYYY-MM format
            df['年月'] = df['月份'].dt.to_period('M').astype(str)
            df['生产年月'] = df['生产月'].dt.to_period('M').astype(str)
            # Create a pivot table to get the count of machines repaired, indexed by repair month and columns by production month
            pivot_table = pd.pivot_table(df, values='生产月', index='年月', columns='生产年月', aggfunc='count', fill_value=0)

            # Convert pivot table to a more readable format for display
            matrix = pivot_table.reset_index()
            matrix.columns.name = None  # Remove the name of the index/columns for cleaner output
            # Add row sums
            row_sums = matrix.iloc[:, 1:].sum(axis=1)

            matrix['报修年月_dt'] = pd.to_datetime(matrix['年月'])

            # Iterate over each cell to check the condition and replace the value if necessary, without altering original date formats
            for col in matrix.columns[1:-1]:  # Exclude the first column which contains the dates and the last temporary datetime column
                col_date = pd.to_datetime(col)
                for index, row in matrix.iterrows():
                    row_date = row['报修年月_dt']
                    if col_date > row_date:
                        matrix.at[index, col] = None

            # Remove the temporary datetime column
            matrix.drop(columns=['报修年月_dt'], inplace=True)
            #插入合计
            matrix.insert(1, '维修合计', row_sums)

        
            output = BytesIO()
            # 将DataFrame写入到BytesIO对象中（作为Excel文件）
            matrix.to_excel(output, index=False, engine='openpyxl')  # 确保使用openpyxl引擎来处理xlsx文件

            # 重要：为了从BytesIO对象读取数据，需要将指针移动到开始位置
            output.seek(0)
        st.success('处理完成!')

        st.download_button(
                    label = "Download data as Excel sheet",
                    data = output,
                    file_name = 'pivot_matrix.xlsx'
                )
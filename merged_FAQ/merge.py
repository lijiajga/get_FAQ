# pip install pandas openpyxl
import os
import pandas as pd

# 1. 指定目录
folder = r'D:\Pan_xfusion_worksapce\1.《《任务》》\8月\4.知识库文件备份\ISC+通用-FAQ库-内部公开\FAQ_all'          # 可改成绝对路径
out_file = r'./merged.xlsx'    # 输出的总表

# 2. 找到所有支持的文件
files = [f for f in os.listdir(folder)
         if f.lower().endswith(('.xlsx', '.xls', '.csv'))]

if not files:
    raise RuntimeError('目录里没有 Excel/CSV 文件！')

# 3. 逐个读入并追加
all_df = []
for f in files:
    full_name = os.path.join(folder, f)
    print(f'正在读取：{f}')
    if f.lower().endswith('.csv'):
        # 如果 CSV 不是 utf-8，试试 encoding='gbk' 或 'utf-8-sig'
        df = pd.read_csv(full_name)
    else:
        # 支持多工作表：把所有 sheet 都读出来再 concat
        sheets = pd.read_excel(full_name, sheet_name=None)   # 字典
        df = pd.concat(sheets.values(), ignore_index=True)
    # 可选：新增一列标识来源文件
    df['来源文件'] = f
    all_df.append(df)

merged = pd.concat(all_df, ignore_index=True)

# 4. 导出
merged.to_excel(out_file, index=False)
print(f'合并完成，共 {len(merged)} 行，结果保存在：{out_file}')

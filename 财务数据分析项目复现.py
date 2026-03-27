import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from functools import reduce
from datetime import datetime, timedelta

# 设置绘图风格和字体（解决中文乱码问题）
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号


#--- 数据读取 ---
all_datasheet = pd.read_excel('pvc_industry_data.xlsx',sheet_name=None)
df_list = all_datasheet.values()
df = reduce(
    lambda left, right: pd.merge(left, right, on='日期', how='outer'),
    df_list)
df.set_index('日期', inplace=True) # 确保日期是索引



# -------- 数据清洗 -------
# 1. 检查缺失值
print("数据缺失情况检查：")
print(df.isnull().sum())

# 展示前5行数据
print("\n数据预览：")
print(df.head())


# ------勾稽关系校验 -------
# 1. 验证毛利计算：销售收入 - 营业成本 = 毛利
df['计算毛利'] = df['销售收入（元）'] - df['营业成本（元）']
df['毛利差异'] = abs(df['毛利（元）'] - df['计算毛利'])
check_gross_profit = (df['毛利差异'] < 0.01).all() # 允许极小的浮点数误差

# 2. 验证现金周期公式：存货周转天数 + 应收账款周转天数 - 应付账款周转天数 = 现金周期
df['计算现金周期'] = df['存货周转天数_y'] + df['应收账款周转天数_y'] - df['应付账款周转天数']
df['现金周期差异'] = abs(df['现金周期（天）'] - df['计算现金周期'])
check_cash_cycle = (df['现金周期差异'] < 0.1).all()

print(f"毛利勾稽关系校验: {'通过' if check_gross_profit else '勾稽数据不一致，请检查原始数据！！！'}")
print(f"现金周期勾稽关系校验: {'通过' if check_cash_cycle else '勾稽数据不一致，请检查原始数据！！！'}")


#------ 经营业绩分析 ------
# 计算关键财务指标
df['毛利率'] = df['毛利（元）'] / df['销售收入（元）']
df['净利率'] = df['净利润（元）'] / df['销售收入（元）']
df['收入环比增长'] = df['销售收入（元）'].pct_change()
df['收入同比增长'] = df['销售收入（元）'].pct_change(periods=12)

# 绘制收入与利润趋势图
fig, ax1 = plt.subplots(figsize=(14, 6))

# 左轴：销售收入
color = 'tab:blue'
ax1.set_xlabel('日期')
ax1.set_ylabel('销售收入 (万元)', color=color)
ax1.bar(df.index, df['销售收入（元）'] / 10000, color=color, alpha=0.6, label='销售收入')
ax1.tick_params(axis='y', labelcolor=color)

# 右轴：毛利率与净利率
ax2 = ax1.twinx()
color = 'tab:red'
ax2.set_ylabel('利润率 (%)', color=color)
ax2.plot(df.index, df['毛利率'], color='red', marker='o', label='毛利率')
ax2.plot(df.index, df['净利率'], color='green', linestyle='--', marker='x', label='净利率')
ax2.tick_params(axis='y', labelcolor=color)
ax2.legend(loc='upper left')

plt.title('销售收入与利润率趋势分析')
plt.savefig('图1_销售收入与利润率趋势.png', dpi=300, bbox_inches='tight')
plt.show()

# 分析结论输出
df['年份'] = df.index.year
def get_max_profit_info(group):
    max_profit_month = df['净利润（元）'].idxmax()
    max_idx = group['净利润（元）'].idxmax()
    return pd.Series({
        '净利润最高月份': max_idx.strftime('%Y-%m-%d'),
        '净利润最高金额(元)': group.loc[max_idx, '净利润（元）'],
        '平均毛利率': group['毛利率'].mean(),
        '平均净利率': group['净利率'].mean()
    })
analysis_summary = df.groupby('年份').apply(get_max_profit_info, include_groups=False)


# ------ 成本驱动因素分析 ------
# 相关性分析
corr_matrix = df[['PVC树脂价格（元/吨）', '营业成本（元）', '单位成本（元/米）', '毛利率']].corr()

# 可视化：PVC价格 vs 毛利率
fig, ax = plt.subplots(figsize=(12, 6))
sns.regplot(x='PVC树脂价格（元/吨）', y='毛利率', data=df, ax=ax, color='purple')
ax.set_title('PVC树脂价格与毛利率的相关性分析')
ax.set_xlabel('PVC树脂价格 (元/吨)')
ax.set_ylabel('毛利率')
plt.grid(True, linestyle='--', alpha=0.5)
plt.savefig('图2_PVC价格与毛利率相关性.png', dpi=300, bbox_inches='tight')
plt.show()

# 成本传导分析
fig, ax1 = plt.subplots(figsize=(14, 6))
ax1.plot(df.index, df['PVC树脂价格（元/吨）'], color='black', label='PVC树脂价格(元/吨)', marker='s')
ax1.set_ylabel('PVC树脂价格 (元/吨）')
ax1.legend(loc='upper left')

ax2 = ax1.twinx()
ax2.plot(df.index, df['营业成本（元）'], color='orange', label='月度营业成本(元)', marker='^')
ax2.set_ylabel('营业成本 (元)')
ax2.legend(loc='upper right')

plt.title('原材料价格与营业成本走势对比')
plt.savefig('图3_原材料与成本走势.png', dpi=300, bbox_inches='tight')
plt.show()


# ------ 运营效率与现金流分析 ------
# 现金周期趋势分析
fig, ax = plt.subplots(figsize=(14, 6))
ax.stackplot(df.index,
             df['存货周转天数_y'],
             df['应收账款周转天数_y'],
             -df['应付账款周转天数'],
             labels=['存货周转天数', '应收账款周转天数', '-应付账款周转天数'],
             colors=['#ff9999','#66b3ff','#99ff99'])
ax.plot(df.index, df['现金周期（天）'], color='black', linewidth=2.5, label='现金周期(净)', marker='o')

ax.axhline(0, color='black', linewidth=1)
ax.set_title('现金周期构成与趋势 (堆叠图)')
ax.set_ylabel('天数')
ax.legend(loc='upper right')
plt.savefig('图4_现金周期分析.png', dpi=300, bbox_inches='tight')
plt.show()

# 运营效率统计
efficiency_stats = pd.DataFrame({
    '统计项': ['平均现金周期(天)', '最长现金周期(天)', '最短现金周期(天)'],
    '数值': [
        round(df['现金周期（天）'].mean(), 2),
        round(df['现金周期（天）'].max(), 2),
        round(df['现金周期（天）'].min(), 2)
    ]
})


# ------杜邦分析------
# 简化的因素分解：净利润 = 销售收入 * (毛利率 - 费用率)
df['费用率'] = (df['销售收入（元）'] - df['毛利（元）'] - df['净利润（元）']) / df['销售收入（元）'] # 这里仅为推算运营费用占比

# 对比不同时期的盈利驱动因素
# 选取第一季度和第四季度的平均数据进行对比
q1_data = df.loc[df.index.quarter == 1, ['毛利率','净利率','费用率']].mean()
q4_data = df.loc[df.index.quarter == 4, ['毛利率','净利率','费用率']].mean()

comparison = pd.DataFrame({
    '指标': ['毛利率', '净利率', '费用率'],
    'Q1均值': [q1_data['毛利率'], q1_data['净利率'], q1_data['费用率']],
    'Q4均值': [q4_data['毛利率'], q4_data['净利率'], q4_data['费用率']]
})


# ------  自动化报告生成逻辑 ------
latest = df.iloc[-1]
prev = df.iloc[-2]

report_lines = []
report_lines.append(f"=========== PVC行业财务简报 ({latest.name.strftime('%Y-%m')}) ===========")
report_lines.append("")
report_lines.append("【经营业绩】")
report_lines.append(f"本月实现销售收入 {latest['销售收入（元）']/10000:.2f} 万元，")
report_lines.append(f"环比 {'增长' if latest['销售收入（元）'] > prev['销售收入（元）'] else '下降'} {abs(latest['收入环比增长'])*100:.2f}%。")
report_lines.append(f"本月净利润为 {latest['净利润（元）']/10000:.2f} 万元，净利率为 {latest['净利率']:.2%}。")
report_lines.append("")
report_lines.append("【成本分析】")
report_lines.append(f"本月PVC树脂价格为 {latest['PVC树脂价格（元/吨）']:.2f} 元/吨。")
report_lines.append(f"毛利率录得 {latest['毛利率']:.2%}，处于历史 {'高位' if latest['毛利率'] > df['毛利率'].median() else '低位'}。")
report_lines.append("")
report_lines.append("【运营效率】")
report_lines.append(f"当前现金周期为 {latest['现金周期（天）']:.2f} 天。")
report_lines.append(f"其中存货周转天数 {latest['存货周转天数_y']:.2f} 天，应收账款周转天数 {latest['应收账款周转天数_y']:.2f} 天。")
report_lines.append(f"资金周转效率 {'提升' if latest['现金周期（天）'] < prev['现金周期（天）'] else '降低'}。")
report_lines.append("")
report_lines.append("【风险提示】")

if latest['净利润（元）'] < 0:
    report_lines.append("警告：本月出现亏损，需重点关注成本控制。")
if latest['现金周期（天）'] > 90:
    report_lines.append("警告：现金周期超过3个月，存在较大的资金链压力。")
if latest['净利润（元）'] >= 0 and latest['现金周期（天）'] <= 90:
    report_lines.append("暂无明显风险预警。")

# 将报告转换为DataFrame以便写入Excel
report_df = pd.DataFrame({'分析报告内容': report_lines})


#------ 导出相关分析数据 ------
output_file = 'PVC行业分析报告结果.xlsx'

print(f"正在将分析结果导出到 {output_file} ...")

with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # 1. 原始数据及计算列
    df.to_excel(writer, sheet_name='完整数据表')

    # 2. 业绩分析结论
    analysis_summary.to_excel(writer, sheet_name='业绩分析汇总', index=False)

    # 3. 相关性矩阵
    corr_matrix.to_excel(writer, sheet_name='相关性矩阵')

    # 4. 运营效率统计
    efficiency_stats.to_excel(writer, sheet_name='运营效率统计', index=False)

    # 5. 季度对比
    comparison.to_excel(writer, sheet_name='季度指标对比', index=False)

    # 8. 文字报告
    report_df.to_excel(writer, sheet_name='自动化报告', index=False)

print("导出完成！")
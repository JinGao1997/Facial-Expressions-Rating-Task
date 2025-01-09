import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# 读取数据文件
data = pd.read_excel('aligned_data.xlsx')  # 确保 'aligned_data.xlsx' 文件路径正确

# 从 Material 列中提取表情类型
data['Expression_Type'] = data['Material'].apply(lambda x: 'Disgust' if 'dis' in x else 
                                                           'Enjoyment' if 'enj' in x else 
                                                           'Affiliation' if 'aff' in x else 
                                                           'Dominance' if 'dom' in x else 
                                                           'Neutral' if 'neu' in x else 'Other')

# 过滤掉 'Other' 类别
filtered_data = data[data['Expression_Type'] != 'Other']

# 计算每个图片的平均 Arousal 和 Realism 评分
mean_scores = filtered_data.groupby(['Material', 'Expression_Type']).agg(
    Arousal_Score=('Arousal_Score', 'mean'),
    Realism_Score=('Realism_Score', 'mean')
).reset_index()

# 更新配色方案
updated_colors = {
    'Enjoyment': '#1F77B4',  # 深蓝色
    'Neutral': '#555555',    # 深灰色（略微加深）
    'Disgust': '#A14D4D',    # 棕红色（提高亮度）
    'Affiliation': '#2CA02C',# 浅绿色（保持不变）
    'Dominance': '#FF7F0E'   # 深橙色（保持不变）
}

# 直接将颜色与类别对应
palette = updated_colors

# 指定绘图时的顺序
order = ['Enjoyment', 'Affiliation', 'Dominance', 'Disgust', 'Neutral']

# 小提琴图绘制
plt.figure(figsize=(14, 6))

# 小提琴图 - Arousal Score
plt.subplot(1, 2, 1)
sns.violinplot(x='Expression_Type', y='Arousal_Score', data=mean_scores, 
               order=order, palette=palette, inner=None)
sns.boxplot(x='Expression_Type', y='Arousal_Score', data=mean_scores, 
            order=order, width=0.1, palette=palette, fliersize=0)
plt.title('Arousal Scores by Expression Type', fontsize=15, fontweight='bold')
plt.xlabel('Expression Type', fontsize=12, fontweight='bold')
plt.ylabel('Arousal Score', fontsize=12, fontweight='bold')
plt.xticks(fontsize=10)
plt.yticks(fontsize=10)
plt.grid(axis='y', linestyle='--', alpha=0.7)

# 小提琴图 - Realism Score
plt.subplot(1, 2, 2)
sns.violinplot(x='Expression_Type', y='Realism_Score', data=mean_scores, 
               order=order, palette=palette, inner=None)
sns.boxplot(x='Expression_Type', y='Realism_Score', data=mean_scores, 
            order=order, width=0.1, palette=palette, fliersize=0)
plt.title('Plausibility Scores by Expression Type', fontsize=15, fontweight='bold')
plt.xlabel('Expression Type', fontsize=12, fontweight='bold')
plt.ylabel('Plausibility Score', fontsize=12, fontweight='bold')
plt.xticks(fontsize=10)
plt.yticks(fontsize=10)
plt.grid(axis='y', linestyle='--', alpha=0.7)

plt.tight_layout()

# 保存并显示图像
plt.savefig('arousal_plausibility_violin_plots_updated.png', dpi=300)
plt.show()

# 条形图绘制
# 将 Expression_Type 映射到 Categorizing_Expressions_Score
mapping = {
    'Enjoyment': 1,
    'Affiliation': 2,
    'Dominance': 3,
    'Disgust': 4,
    'Neutral': 5
}
data['Correct'] = data.apply(lambda row: 1 if mapping[row['Expression_Type']] == row['Categorizing_Expressions_Score'] else 0, axis=1)

# 计算 Percentage Hit Rate
percentage_hit_rate = data.groupby(['Material', 'Expression_Type']).agg(Hit_Rate=('Correct', 'mean')).reset_index()

# 计算描述性统计信息
hit_rate_summary = percentage_hit_rate.groupby('Expression_Type').agg(
    Mean_Hit_Rate=('Hit_Rate', 'mean'),
    SD_Hit_Rate=('Hit_Rate', 'std'),
    N=('Hit_Rate', 'count')
).reset_index()

# 打印计算出的每个情绪类别的命中率和标准差
print("Mean Hit Rate and Standard Deviation by Expression Type:")
print(hit_rate_summary)

# 绘制条形图
color_list = [updated_colors.get(x, '#FFCD9B') for x in hit_rate_summary['Expression_Type']]

plt.figure(figsize=(10, 6))
sns.barplot(x='Expression_Type', y='Mean_Hit_Rate', data=hit_rate_summary, palette=color_list, ci=None)

# 添加误差条
for i in range(len(hit_rate_summary)):
    plt.errorbar(hit_rate_summary['Expression_Type'][i], 
                 hit_rate_summary['Mean_Hit_Rate'][i], 
                 yerr=hit_rate_summary['SD_Hit_Rate'][i], 
                 fmt='none', c='black', capsize=5)

# 设置图表标题和标签
plt.title('Mean Percentage Hit Rate by Expression Type', fontsize=15, fontweight='bold')
plt.xlabel('Expression Type', fontsize=12, fontweight='bold')
plt.ylabel('Mean Percentage Hit Rate', fontsize=12, fontweight='bold')
plt.ylim(0, 1)
plt.xticks(fontsize=10)
plt.yticks(fontsize=10)
plt.grid(axis='y', linestyle='--', alpha=0.7)
plt.tight_layout()

# 保存图形
output_bar_image_path_custom = 'percentage_hit_rate_plot_updated.png'
plt.savefig(output_bar_image_path_custom, dpi=300)

# 显示图形
plt.show()

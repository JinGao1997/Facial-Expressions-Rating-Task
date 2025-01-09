import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# 动态生成星号
def get_stars(p):
    if p < 0.001:
        return "***"
    elif p < 0.01:
        return "**"
    elif p < 0.05:
        return "*"
    else:
        return "ns"

# 标注函数
def add_stat_annotation(ax, x1, x2, y, stars, line_height=0.03, star_offset=0.015, color='gray'):
    """
    添加显著性标注线和星号。
    x1, x2: 两组的横坐标
    y: 线的起始高度
    stars: 星号
    line_height: 连线高度
    star_offset: 星号与连线之间的垂直距离
    """
    ax.plot([x1, x1, x2, x2], [y, y + line_height, y + line_height, y], color=color, lw=0.8, linestyle='--')
    ax.text((x1 + x2) / 2, y + line_height + star_offset, stars, ha='center', va='bottom', color='black', fontsize=10)

# 数据和配色方案
data = pd.read_excel('aligned_data.xlsx')
data['Expression_Type'] = data['Material'].apply(lambda x: 'Disgust' if 'dis' in x else 
                                                 'Enjoyment' if 'enj' in x else 
                                                 'Affiliation' if 'aff' in x else 
                                                 'Dominance' if 'dom' in x else 
                                                 'Neutral' if 'neu' in x else 'Other')
filtered_data = data[data['Expression_Type'] != 'Other']
mean_scores = filtered_data.groupby(['Material', 'Expression_Type']).agg(
    Arousal_Score=('Arousal_Score', 'mean'),
    Realism_Score=('Realism_Score', 'mean')
).reset_index()
palette = {'Enjoyment': '#1F77B4', 'Neutral': '#555555', 'Disgust': '#A14D4D', 
           'Affiliation': '#2CA02C', 'Dominance': '#FF7F0E'}
order = ['Enjoyment', 'Affiliation', 'Dominance', 'Disgust', 'Neutral']

arousal_results = pd.DataFrame({
    "Comparison": [
        "Affiliation - Disgust", "Affiliation - Dominance", "Disgust - Dominance",
        "Affiliation - Enjoyment", "Disgust - Enjoyment", "Dominance - Enjoyment",
        "Affiliation - Neutral", "Disgust - Neutral", "Dominance - Neutral", "Enjoyment - Neutral"
    ],
    "P.adj": [
        5.025533e-15, 6.408629e-11, 1.000000e+00,
        2.361947e-67, 7.926721e-20, 2.940403e-25,
        1.307385e-15, 2.497346e-59, 8.407757e-51, 3.723925e-145
    ]
})
arousal_results["Stars"] = arousal_results["P.adj"].apply(get_stars)

realism_results = pd.DataFrame({
    "Comparison": [
        "Affiliation - Disgust", "Affiliation - Dominance", "Disgust - Dominance",
        "Affiliation - Enjoyment", "Disgust - Enjoyment", "Dominance - Enjoyment",
        "Affiliation - Neutral", "Disgust - Neutral", "Dominance - Neutral", "Enjoyment - Neutral"
    ],
    "P.adj": [
        9.493641e-12, 2.100988e-16, 1.000000e+00,
        1.914493e-07, 2.841963e-37, 3.408678e-45,
        8.308490e-18, 1.419330e-57, 2.194509e-67, 1.211131e-02
    ]
})
realism_results["Stars"] = realism_results["P.adj"].apply(get_stars)

# 绘图
plt.figure(figsize=(16, 6))

# Arousal Score
ax1 = plt.subplot(1, 2, 1)
sns.violinplot(x='Expression_Type', y='Arousal_Score', data=mean_scores, order=order, palette=palette, inner=None)
sns.boxplot(x='Expression_Type', y='Arousal_Score', data=mean_scores, order=order, width=0.1, palette=palette, fliersize=0)
plt.title('Arousal Scores by Expression Type', fontsize=15, fontweight='bold')
plt.xlabel('Expression Type', fontsize=12, fontweight='bold')
plt.ylabel('Arousal Score', fontsize=12, fontweight='bold')
plt.ylim(1, 9)
y_max = 9
for i, row in arousal_results.iterrows():
    group1, group2 = row["Comparison"].split(" - ")
    if group1 in order and group2 in order:
        x1, x2 = order.index(group1), order.index(group2)
        add_stat_annotation(ax1, x1, x2, y=y_max - 0.3 - i * 0.3, stars=row["Stars"], star_offset=0.0001)

# Realism Score
ax2 = plt.subplot(1, 2, 2)
sns.violinplot(x='Expression_Type', y='Realism_Score', data=mean_scores, order=order, palette=palette, inner=None)
sns.boxplot(x='Expression_Type', y='Realism_Score', data=mean_scores, order=order, width=0.1, palette=palette, fliersize=0)
plt.title('Plausibility Scores by Expression Type', fontsize=15, fontweight='bold')
plt.xlabel('Expression Type', fontsize=12, fontweight='bold')
plt.ylabel('Plausibility Score', fontsize=12, fontweight='bold')
plt.ylim(1, 7)
y_max = 7
for i, row in realism_results.iterrows():
    group1, group2 = row["Comparison"].split(" - ")
    if group1 in order and group2 in order:
        x1, x2 = order.index(group1), order.index(group2)
        add_stat_annotation(ax2, x1, x2, y=y_max - 0.3 - i * 0.3, stars=row["Stars"], star_offset=0.0001)

plt.tight_layout()
plt.savefig('adjusted_violin_plots_with_closer_stars.png', dpi=300)
plt.show()

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

# 读取数据
file_path = "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data = pd.read_excel(file_path)

# 定义目标 Expressors 列表
target_female_expressors = ["Fema32", "Fema46", "Fema64", "Fema16", "Fema30", 
                            "Fema10", "Fema66", "Fema24", "Fema4", "Fema12", 
                            "Fema60", "Fema82", "Fema6", "Fema68", "Fema86", 
                            "Fema56", "Fema38", "Fema88", "Fema80", "Fema28", 
                            "Fema58", "Fema20", "Fema26", "Fema52", "Fema50", 
                            "Fema40", "Fema22", "Fema8", "Fema78", "Fema14"]

target_male_expressors = ["Male29", "Male37", "Male73", "Male45", "Male35", 
                          "Male5", "Male53", "Male63", "Male47", "Male81", 
                          "Male49", "Male21", "Male79", "Male67", "Male69", 
                          "Male17", "Male15", "Male65", "Male39", "Male31", 
                          "Male51", "Male59", "Male23", "Male85", "Male61", 
                          "Male71", "Male3", "Male19", "Male27", "Male33"]

# 从 Material 列中提取 Expressor, 性别 和情绪类型信息
data['Expression_Type'] = data['Material'].apply(lambda x: 
    'Disgust' if 'dis' in x else
    'Enjoyment' if 'enj' in x else
    'Affiliation' if 'aff' in x else
    'Dominance' if 'dom' in x else
    'Neutral' if 'neu' in x else 'Other'
)

data['Expressor'] = data['Material'].str.extract(r'(Fema\d+|Male\d+)')
data['Gender'] = data['Expressor'].apply(lambda x: 'Female' if 'Fema' in x else 'Male')

# 创建 Expressor_Short 列并确保其顺序与目标顺序一致
data['Expressor_Short'] = pd.Categorical(data['Expressor'].str.replace('Fema', 'F').str.replace('Male', 'M'), 
    categories=[x.replace('Fema', 'F').replace('Male', 'M') for x in (target_female_expressors + target_male_expressors)],
    ordered=True
)

# 过滤数据
data = data[data['Expressor'].isin(target_female_expressors + target_male_expressors)]

# 计算 Percentage Hit Rate 和 Realism 平均得分
expression_mapping = {'Enjoyment': 1, 'Affiliation': 2, 'Dominance': 3, 'Disgust': 4, 'Neutral': 5}

data['Correct'] = np.where(data['Categorizing_Expressions_Score'] == 6, np.nan, 
                           np.where(data['Categorizing_Expressions_Score'] == data['Expression_Type'].map(expression_mapping), 1, 0))

summary_data = data.groupby(['Expressor_Short', 'Gender']).agg(
    Hit_Rate=('Correct', 'mean'),
    Avg_Realism=('Realism_Score', 'mean')
).reset_index()

# 计算 UHR (Unbiased Hit Rate)
uhr_data = data.pivot_table(index=['Expressor_Short', 'Gender', 'Expression_Type'], 
                            columns='Categorizing_Expressions_Score', 
                            values='Material', 
                            aggfunc='count', 
                            fill_value=0).reset_index()

uhr_data['row_sum'] = uhr_data.iloc[:, 3:].sum(axis=1)

def calc_uhr(row):
    score_col = expression_mapping.get(row['Expression_Type'], np.nan)
    if pd.notna(score_col):
        correct_preds = row[score_col]
        if correct_preds > 0:
            return (correct_preds / row['row_sum']) * (correct_preds / row['row_sum'])
        return 0
    return 0

uhr_data['UHR'] = uhr_data.apply(calc_uhr, axis=1)
uhr_summary = uhr_data.groupby(['Expressor_Short', 'Gender']).agg(Average_UHR=('UHR', 'mean')).reset_index()

# 合并所有结果
final_summary = pd.merge(summary_data, uhr_summary, on=['Expressor_Short', 'Gender'])

# 检查 final_summary 内容
print("Final Summary Head:")
print(final_summary.head())

# 保存处理后的数据以供可视化使用
final_summary.to_csv("final_summary_data.csv", index=False)

# 可视化部分（按三行一列排列，并调整图例位置）
def plot_combined_data(df):
    sns.set(style="whitegrid")
    fig, axes = plt.subplots(3, 1, figsize=(16, 18))  # 3行1列的图表
    
    # 绘制 Percentage Hit Rate
    sns.barplot(x='Expressor_Short', y='Hit_Rate', hue='Gender', data=df, ax=axes[0], palette="Blues_d")
    axes[0].set_title("Percentage Hit Rate", fontsize=18)
    axes[0].set_xlabel("Expressor", fontsize=14)
    axes[0].set_ylabel("Percentage Hit Rate", fontsize=14)
    axes[0].tick_params(axis='x', rotation=90, labelsize=12)
    axes[0].tick_params(axis='y', labelsize=14)
    axes[0].legend(loc='upper right', bbox_to_anchor=(1.15, 1))  # 调整图例位置

    # 绘制 UHR
    sns.barplot(x='Expressor_Short', y='Average_UHR', hue='Gender', data=df, ax=axes[1], palette="Oranges_d")
    axes[1].set_title("Average UHR", fontsize=18)
    axes[1].set_xlabel("Expressor", fontsize=14)
    axes[1].set_ylabel("Average UHR", fontsize=14)
    axes[1].tick_params(axis='x', rotation=90, labelsize=12)
    axes[1].tick_params(axis='y', labelsize=14)
    axes[1].legend(loc='upper right', bbox_to_anchor=(1.15, 1))  # 调整图例位置

    # 绘制 Average Realism
    sns.barplot(x='Expressor_Short', y='Avg_Realism', hue='Gender', data=df, ax=axes[2], palette="Greens_d")
    axes[2].set_title("Average Realism Score", fontsize=18)
    axes[2].set_xlabel("Expressor", fontsize=14)
    axes[2].set_ylabel("Average Realism Score", fontsize=14)
    axes[2].tick_params(axis='x', rotation=90, labelsize=12)
    axes[2].tick_params(axis='y', labelsize=14)
    axes[2].legend(loc='upper right', bbox_to_anchor=(1.15, 1))  # 调整图例位置

    plt.tight_layout(pad=3.0, rect=[0, 0, 1, 0.98])  # 调整布局，避免标题和标签重叠
    plt.savefig("combined_summary_plot.png", dpi=300)
    plt.show()

# 绘制图表
plot_combined_data(final_summary)


import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

# 读取数据
file_path = "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data = pd.read_excel(file_path)

# 从 Material 列中提取 Expressor, 性别 和情绪类型信息
data['Expression_Type'] = data['Material'].apply(lambda x: 
    'Disgust' if 'dis' in x else
    'Enjoyment' if 'enj' in x else
    'Affiliation' if 'aff' in x else
    'Dominance' if 'dom' in x else
    'Neutral' if 'neu' in x else 'Other'
)

data['Expressor'] = data['Material'].str.extract(r'(Fema\d+|Male\d+)')
data['Gender'] = data['Expressor'].apply(lambda x: 'Female' if 'Fema' in x else 'Male')
data['Expressor_Short'] = data['Expressor'].str.replace('Fema', 'F').str.replace('Male', 'M')

# 筛选出你关注的Expressors，并按照你提供的顺序排列
target_female_expressors = ["Fema32", "Fema46", "Fema64", "Fema16", "Fema30", 
                            "Fema10", "Fema66", "Fema24", "Fema4", "Fema12", 
                            "Fema60", "Fema82", "Fema6", "Fema68", "Fema86", 
                            "Fema56", "Fema38", "Fema88", "Fema80", "Fema28", 
                            "Fema58", "Fema20", "Fema26", "Fema52", "Fema50", 
                            "Fema40", "Fema22", "Fema8", "Fema78", "Fema14"]

target_male_expressors = ["Male29", "Male37", "Male73", "Male45", "Male35", 
                          "Male5", "Male53", "Male63", "Male47", "Male81", 
                          "Male49", "Male21", "Male79", "Male67", "Male69", 
                          "Male17", "Male15", "Male65", "Male39", "Male31", 
                          "Male51", "Male59", "Male23", "Male85", "Male61", 
                          "Male71", "Male3", "Male19", "Male27", "Male33"]

# 过滤数据
data = data[data['Expressor'].isin(target_female_expressors + target_male_expressors)]

# 按情绪类型计算 Arousal 的平均分
arousal_data = data.groupby(['Expressor_Short', 'Expression_Type']).agg(
    Avg_Arousal=('Arousal_Score', 'mean')
).reset_index()

# 创建雷达图函数
def create_radar_chart(ax, df, expressor_name):
    # 使用缩写形式的情绪标签
    categories = ['Enj', 'Aff', 'Dom', 'Dis', 'Neu']
    values = df.set_index('Expression_Type').reindex(
        ['Enjoyment', 'Affiliation', 'Dominance', 'Disgust', 'Neutral']
    )['Avg_Arousal'].values
    
    angles = np.linspace(0, 2 * np.pi, len(categories), endpoint=False).tolist()
    values = np.concatenate((values, [values[0]]))
    angles += angles[:1]
    
    ax.fill(angles, values, color='#FFA76D', alpha=0.25)
    ax.plot(angles, values, color='#FFA76D', linewidth=2)
    ax.set_yticklabels([])

    # 设置标签并标注得分范围
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(categories, fontsize=10)
    ax.set_ylim(1, 9)  # 设置Arousal分数范围
    ax.set_yticks([1, 3, 5, 7, 9])
    ax.set_yticklabels(["1", "3", "5", "7", "9"], fontsize=8)
    ax.set_title(expressor_name, size=12, color='#C76B24', y=1.1)

# 为女性 Expressors 绘制雷达图
num_female = len(target_female_expressors)
cols = 5  # 每行显示5个图
rows = int(np.ceil(num_female / cols))

fig, axes = plt.subplots(rows, cols, figsize=(20, 4 * rows), subplot_kw=dict(polar=True))
fig.suptitle("Female Expressors' Arousal Scores", fontsize=20, fontweight='bold')

for i, expressor in enumerate(target_female_expressors):
    row, col = divmod(i, cols)
    short_name = expressor.replace('Fema', 'F')
    expressor_data = arousal_data[arousal_data['Expressor_Short'] == short_name]
    create_radar_chart(axes[row, col], expressor_data, short_name)

# 删除空白子图
for i in range(num_female, rows * cols):
    fig.delaxes(axes.flatten()[i])

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("combined_radar_female.png", dpi=300)
plt.show()

# 为男性 Expressors 绘制雷达图
num_male = len(target_male_expressors)
cols = 5  # 每行显示5个图
rows = int(np.ceil(num_male / cols))

fig, axes = plt.subplots(rows, cols, figsize=(20, 4 * rows), subplot_kw=dict(polar=True))
fig.suptitle("Male Expressors' Arousal Scores", fontsize=20, fontweight='bold')

for i, expressor in enumerate(target_male_expressors):
    row, col = divmod(i, cols)
    short_name = expressor.replace('Male', 'M')
    expressor_data = arousal_data[arousal_data['Expressor_Short'] == short_name]
    create_radar_chart(axes[row, col], expressor_data, short_name)

# 删除空白子图
for i in range(num_male, rows * cols):
    fig.delaxes(axes.flatten()[i])

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("combined_radar_male.png", dpi=300)
plt.show()

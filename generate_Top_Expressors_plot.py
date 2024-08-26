import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
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

data = data[data['Expressor'].isin(target_female_expressors + target_male_expressors)]

# 计算 Percentage Hit Rate, Arousal 和 Realism 平均得分
expression_mapping = {'Enjoyment': 1, 'Affiliation': 2, 'Dominance': 3, 'Disgust': 4, 'Neutral': 5}

data['Correct'] = np.where(data['Categorizing_Expressions_Score'] == 6, np.nan, 
                           np.where(data['Categorizing_Expressions_Score'] == data['Expression_Type'].map(expression_mapping), 1, 0))

summary_data = data.groupby(['Expressor_Short', 'Gender']).agg(
    Hit_Rate=('Correct', 'mean'),
    Avg_Arousal=('Arousal_Score', 'mean'),
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



# 自定义配色 (基于橘色调)
custom_palette_1 = sns.light_palette("#D98A3D", n_colors=1, input="hex", reverse=True)
custom_palette_2 = sns.light_palette("#C76B24", n_colors=1, input="hex", reverse=True)
custom_palette_3 = sns.light_palette("#E17B41", n_colors=1, input="hex", reverse=True)
custom_palette_4 = sns.light_palette("#FFA76D", n_colors=1, input="hex", reverse=True)
custom_palette_5 = sns.light_palette("#FFB383", n_colors=1, input="hex", reverse=True)

# 按性别区分可视化
def plot_gender_data(df, gender, title_prefix, expressors):
    df = df[df['Expressor_Short'].isin(expressors)]
    df['Expressor_Short'] = pd.Categorical(df['Expressor_Short'], categories=expressors, ordered=True)
    
    sns.set(style="whitegrid")
    fig, axes = plt.subplots(2, 2, figsize=(18, 16))  # 增加图表尺寸
    fig.suptitle(f"{title_prefix} Expressors", fontsize=24, fontweight='bold')

    sns.barplot(x='Expressor_Short', y='Hit_Rate', data=df, ax=axes[0, 0], palette=custom_palette_1)
    axes[0, 0].set_title(f"Percentage Hit Rate", fontsize=18)
    axes[0, 0].set_xlabel("Expressor", fontsize=14)
    axes[0, 0].set_ylabel("Percentage Hit Rate", fontsize=14)
    axes[0, 0].tick_params(axis='x', rotation=90, labelsize=12)  # 调整横坐标标签字号
    axes[0, 0].tick_params(axis='y', labelsize=14)  # 调整纵坐标标签字号

    sns.barplot(x='Expressor_Short', y='Avg_Arousal', data=df, ax=axes[0, 1], palette=custom_palette_2)
    axes[0, 1].set_title(f"Average Arousal Score", fontsize=18)
    axes[0, 1].set_xlabel("Expressor", fontsize=14)
    axes[0, 1].set_ylabel("Average Arousal Score", fontsize=14)
    axes[0, 1].tick_params(axis='x', rotation=90, labelsize=12)  # 调整横坐标标签字号
    axes[0, 1].tick_params(axis='y', labelsize=14)  # 调整纵坐标标签字号

    sns.barplot(x='Expressor_Short', y='Avg_Realism', data=df, ax=axes[1, 0], palette=custom_palette_3)
    axes[1, 0].set_title(f"Average Realism Score", fontsize=18)
    axes[1, 0].set_xlabel("Expressor", fontsize=14)
    axes[1, 0].set_ylabel("Average Realism Score", fontsize=14)
    axes[1, 0].tick_params(axis='x', rotation=90, labelsize=12)  # 调整横坐标标签字号
    axes[1, 0].tick_params(axis='y', labelsize=14)  # 调整纵坐标标签字号

    sns.barplot(x='Expressor_Short', y='Average_UHR', data=df, ax=axes[1, 1], palette=custom_palette_4)
    axes[1, 1].set_title(f"Average UHR", fontsize=18)
    axes[1, 1].set_xlabel("Expressor", fontsize=14)
    axes[1, 1].set_ylabel("Average UHR", fontsize=14)
    axes[1, 1].tick_params(axis='x', rotation=90, labelsize=12)  # 调整横坐标标签字号
    axes[1, 1].tick_params(axis='y', labelsize=14)  # 调整纵坐标标签字号

    plt.tight_layout(pad=3.5, rect=[0, 0, 1, 0.96])  # 增加子图间距，防止标题和标签重叠
    plt.savefig(f"combined_plots_{gender.lower()}.png", dpi=300)
    plt.show()

# Female Expressors
female_data = final_summary[final_summary['Gender'] == 'Female']
plot_gender_data(female_data, 'Female', 'Female', [x.replace('Fema', 'F') for x in target_female_expressors])

# Male Expressors
male_data = final_summary[final_summary['Gender'] == 'Male']
plot_gender_data(male_data, 'Male', 'Male', [x.replace('Male', 'M') for x in target_male_expressors])

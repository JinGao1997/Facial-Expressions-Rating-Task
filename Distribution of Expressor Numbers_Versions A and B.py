import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D
from matplotlib.patches import Patch
import numpy as np

# 设置 Seaborn 样式（可选）
import seaborn as sns
sns.set(style="whitegrid")

# 加载 Excel 文件
file_path = r'C:\Users\neuro-lab\OneDrive\桌面\E1_UG_GJ\Trials_E1_Serpentine_LR_ClassicOffers.xlsx'
try:
    excel_data = pd.ExcelFile(file_path)
    print("文件读取成功！")
except Exception as e:
    print(f"读取文件时出错：{e}")
    exit()

# 性别颜色映射（用于标记边框颜色）
gender_edgecolor_map = {'Female': 'purple', 'Male': 'green'}

# 初始化数据存储
all_settings_data = []

# 读取和处理数据
for i in range(1, 6):
    sheet_a = f'Setting{i}_A'
    sheet_b = f'Setting{i}_B'
    try:
        sheet_a_df = pd.read_excel(excel_data, sheet_name=sheet_a)
        sheet_b_df = pd.read_excel(excel_data, sheet_name=sheet_b)
    except Exception as e:
        print(f"读取工作表 {sheet_a} 或 {sheet_b} 时出错：{e}")
        continue

    for sheet_df, sheet_name in zip([sheet_a_df, sheet_b_df], [sheet_a, sheet_b]):
        setting_id = sheet_name.split('Setting')[1]
        sti_column = f'Sti_Setting{setting_id}'

        # 确认列存在
        if sti_column not in sheet_df.columns:
            print(f"列 {sti_column} 在工作表 {sheet_name} 中不存在。")
            continue

        # 提取信息
        sheet_df['Expressor_Number'] = sheet_df[sti_column].str.extract(r'(\d+)$').astype(int)
        sheet_df['Direction'] = sheet_df[sti_column].str.extract(r'^(L|R)', expand=False).map({'L': 'Left', 'R': 'Right'})
        # 使用不区分大小写的正则表达式提取 Gender
        sheet_df['Emotion'] = sheet_df[sti_column].str.extract(r'(aff|enj|dis|neu|dom)', expand=False)
        sheet_df['Gender'] = sheet_df[sti_column].str.extract(r'(?i)(fema|female|male)', expand=False).str.capitalize()
        sheet_df['Setting'] = f'Setting{i}'
        sheet_df['Version'] = 'A' if '_A' in sheet_name else 'B'

        # 添加唯一标识符
        sheet_df['Expressor_ID'] = sheet_df['Expressor_Number'].astype(str) + '_' + sheet_df['Setting'] + '_' + sheet_df['Version']
        required_columns = ['Expressor_ID', 'Expressor_Number', 'Direction', 'Emotion', 'Gender', 'Setting', 'Version']
        sheet_df = sheet_df[required_columns]

        all_settings_data.append(sheet_df)

# 合并所有数据
all_settings_df = pd.concat(all_settings_data, ignore_index=True)

# 映射情绪到描述性名称
emotion_mapping = {'aff': 'Affiliation', 'enj': 'Enjoyment', 'dis': 'Disgust', 'neu': 'Neutral', 'dom': 'Dominance'}
all_settings_df['Emotion_Name'] = all_settings_df['Emotion'].map(emotion_mapping)

# 确保 'Gender' 列中的值正确，并处理缺失或未映射的值
all_settings_df['Gender'] = all_settings_df['Gender'].replace({'Fema': 'Female', 'Female': 'Female', 'Male': 'Male'})
all_settings_df['Gender'] = all_settings_df['Gender'].fillna('Unknown')  # 填充缺失的性别为 'Unknown'

# 调试步骤：检查 'Gender' 列中的唯一值
unique_genders = all_settings_df['Gender'].unique()
print(f"\n'Gender' 列中的唯一值: {unique_genders}")

# 检查是否有未被映射的 Gender 值
mapped_genders = list(gender_edgecolor_map.keys())
unmapped_genders = [gender for gender in unique_genders if gender not in mapped_genders and pd.notna(gender)]

print(f"\n未被映射的 Gender 值: {unmapped_genders}")

if len(unmapped_genders) > 0:
    print("未被映射的 Gender 数据点如下：")
    print(all_settings_df[~all_settings_df['Gender'].isin(mapped_genders) & all_settings_df['Gender'].notna()])

    # 为未映射的 Gender 值指定一个默认颜色，例如 'black'
    for gender in unmapped_genders:
        gender_edgecolor_map[gender] = 'black'

# 创建综合标签，用于 X 轴分组（简化为情绪和版本）
all_settings_df['Group_Label'] = all_settings_df['Emotion_Name'] + '_' + all_settings_df['Version']

# 定义情绪颜色映射（根据需要调整颜色）
emotion_palette = {
    'Affiliation': 'skyblue',
    'Enjoyment': 'orange',
    'Disgust': 'red',
    'Neutral': 'grey',
    'Dominance': 'blue'
}

# 定义每个 Setting 的 x 轴偏移量
setting_offsets = {
    'Setting1': -0.15,
    'Setting2': -0.075,
    'Setting3': 0.0,
    'Setting4': 0.075,
    'Setting5': 0.15
}

# 确保 Group_Label 按 Setting 排序
all_settings_df['Setting_Order'] = all_settings_df['Setting'].apply(lambda x: int(x.replace('Setting', '')))
all_settings_df = all_settings_df.sort_values(['Setting_Order', 'Emotion_Name', 'Version'])

# 获取排序后的 Group_Label
group_labels = all_settings_df['Group_Label'].unique()
all_settings_df['Group_Label'] = pd.Categorical(all_settings_df['Group_Label'], categories=group_labels, ordered=True)

# 计算每个 Group_Label 的基准 x 位置
group_label_to_x = {label: idx for idx, label in enumerate(group_labels)}
all_settings_df['Base_X'] = all_settings_df['Group_Label'].map(group_label_to_x)

# 计算每个数据点的实际 x 位置（基准 x + Setting 的偏移量）
all_settings_df['Actual_X'] = all_settings_df.apply(lambda row: row['Base_X'] + setting_offsets.get(row['Setting'], 0), axis=1)

# 设置边框颜色
all_settings_df['Edge_Color'] = all_settings_df['Gender'].map(gender_edgecolor_map).fillna('black')

# 检查是否有仍然为 NaN 的 Edge_Color
nan_edgecolors = all_settings_df['Edge_Color'].isna().sum()
if nan_edgecolors > 0:
    print(f"警告：存在 {nan_edgecolors} 个 Edge_Color 为 NaN 的数据点，将使用默认颜色 'black'。")
    all_settings_df['Edge_Color'] = all_settings_df['Edge_Color'].fillna('black')

# 绘图
plt.figure(figsize=(25, 14))  # 增加图表宽度和高度以适应更多数据
ax = plt.gca()

# 获取x轴的类别位置
x_categories = list(all_settings_df['Group_Label'].cat.categories)
x_positions = range(len(x_categories))

# 绘制背景灰色区域以区分不同的 Settings
for idx, label in enumerate(x_categories):
    plt.axvspan(idx - 0.5 - 0.15, idx + 0.5 + 0.15, color='lightgrey', alpha=0.3)

# 使用 Matplotlib 的 scatter 绘制散点图
# 首先根据 'Version' 分组
versions = all_settings_df['Version'].unique()

for version in versions:
    version_df = all_settings_df[all_settings_df['Version'] == version]
    marker_shape = 'X' if version == 'B' else 'o'
    
    scatter = ax.scatter(
        version_df['Actual_X'],
        version_df['Expressor_Number'],
        c=version_df['Emotion_Name'].map(emotion_palette),
        marker=marker_shape,
        s=30,  # 调整标记尺寸
        alpha=0.7,
        edgecolors=version_df['Edge_Color'],
        linewidth=0.5,
        label=version
    )

# 设置 X 轴标签和刻度
plt.xticks(ticks=x_positions, labels=x_categories, rotation=45, ha='right', fontsize=12)
plt.xlabel('Emotion and Version', fontsize=14)

# 添加表达者编号标签，并根据性别着色，偏移位置避免覆盖
for _, row in all_settings_df.iterrows():
    plt.text(
        x=row['Actual_X'],
        y=row['Expressor_Number'] + 0.3,  # 增加偏移量，确保标签在数据点上方
        s=str(row['Expressor_Number']),
        color=row['Edge_Color'],  # 根据性别着色
        ha='center',
        fontsize=7,
        alpha=0.85
    )

# 设置标题和轴标签
plt.title('Distribution of Expressor Numbers in Versions A and B under Different Emotions (by Setting)', fontsize=20)
plt.ylabel('Expressor Number', fontsize=14)

# 创建图例
# 情绪图例
emotion_handles = [Patch(color=color, label=emotion) for emotion, color in emotion_palette.items()]

# 版本图例（使用 Line2D）
version_handles = [
    Line2D([0], [0], marker='o', color='w', label='A', markerfacecolor='black', markersize=8, markeredgecolor='w'),
    Line2D([0], [0], marker='X', color='w', label='B', markerfacecolor='black', markersize=8, markeredgecolor='w')
]

# Setting 图例（使用 Patch）
setting_patches = [Patch(facecolor='lightgrey', edgecolor='lightgrey', label=f'Setting{i}') for i in range(1, 6)]

# 添加情绪图例
first_legend = plt.legend(handles=emotion_handles, title='Emotion', bbox_to_anchor=(1.05, 1), loc='upper left')

# 添加版本图例
plt.legend(handles=version_handles, title='Version', bbox_to_anchor=(1.05, 0.85), loc='upper left')

# 添加 Setting 图例
plt.legend(handles=setting_patches, title='Setting', bbox_to_anchor=(1.05, 0.7), loc='upper left')

# 添加第一个图例（情绪）回到图表中
plt.gca().add_artist(first_legend)

# 设置 y 轴限制（可选）
plt.ylim(bottom=0, top=all_settings_df['Expressor_Number'].max() + 5)

plt.tight_layout()
plt.show()

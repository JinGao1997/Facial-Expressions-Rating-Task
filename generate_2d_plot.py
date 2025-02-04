import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# 读取数据
file_path = 'aligned_data.xlsx'  # 请替换成你的实际文件路径
data = pd.read_excel(file_path)

# 根据Material列提取意图表达
def get_intended_expression(material):
    if 'enj' in material:
        return 'Enjoyment'
    elif 'aff' in material:
        return 'Affiliation'
    elif 'dom' in material:
        return 'Dominance'
    elif 'dis' in material:
        return 'Disgust'
    elif 'neu' in material:
        return 'Neutral'
    else:
        return 'Other'

data['Intended_Expression'] = data['Material'].apply(get_intended_expression)

# 将评分映射到标签
mapping = {1: 'Enjoyment', 2: 'Affiliation', 3: 'Dominance', 4: 'Disgust', 5: 'Neutral', 6: 'Other'}
data['Chosen_Expression'] = data['Categorizing_Expressions_Score'].map(mapping)

# 指定行和列的顺序
intended_order = ['Neutral', 'Enjoyment', 'Disgust', 'Affiliation', 'Dominance']
chosen_order = ['Neutral', 'Enjoyment', 'Disgust', 'Affiliation', 'Dominance', 'Other']

# 生成混淆矩阵（按行归一化后乘以100表示百分比）
confusion_matrix = pd.crosstab(data['Intended_Expression'], data['Chosen_Expression'], normalize='index') * 100

# 确保所有列和行都按指定顺序存在
for col in chosen_order:
    if col not in confusion_matrix.columns:
        confusion_matrix[col] = 0
confusion_matrix = confusion_matrix[chosen_order]
confusion_matrix = confusion_matrix.reindex(intended_order)

# 将混淆矩阵保存为文件（可选）
confusion_matrix.to_csv('confusion_matrix_HitRate.csv')
confusion_matrix.to_excel('confusion_matrix_HitRate.xlsx')

# ----------------------------
# 绘制二维热图
# ----------------------------
plt.figure(figsize=(10, 8))
# 使用Seaborn的heatmap函数，其中annot用于在每个单元格中标注数值，fmt设置数字格式
sns.heatmap(confusion_matrix, annot=True, fmt=".1f", cmap="viridis", cbar=True, linewidths=0.5,
            linecolor='gray', annot_kws={"size":10})

# 设置坐标轴标签和标题
plt.xlabel('Chosen Expression', fontsize=12, labelpad=10)
plt.ylabel('Intended Expression', fontsize=12, labelpad=10)
plt.title('Percentage of Chosen Emotions per Intended Expression', fontsize=14, pad=15)

# 调整刻度标签方向（根据需要）
plt.xticks(rotation=45, ha='right', fontsize=10)
plt.yticks(fontsize=10)

# 保存图像（可选）
plt.savefig('confusion_matrix_2d_heatmap.png', dpi=300, bbox_inches='tight')

# 显示图像
plt.show()

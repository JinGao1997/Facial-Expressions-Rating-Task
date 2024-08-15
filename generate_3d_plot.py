import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

# 读取数据
file_path = 'aligned_data.xlsx'
data = pd.read_excel(file_path)

# 从Material列中提取intended expressions
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

# 映射chosen expressions
mapping = {1: 'enj', 2: 'aff', 3: 'dom', 4: 'dis', 5: 'neu', 6: 'other'}
data['Chosen_Expression'] = data['Categorizing_Expressions_Score'].map(mapping)

# 生成混淆矩阵
confusion_matrix = pd.crosstab(data['Intended_Expression'], data['Chosen_Expression'], normalize='index') * 100

# 准备3D绘图数据
x_labels = confusion_matrix.columns
y_labels = confusion_matrix.index
x, y = np.meshgrid(np.arange(len(x_labels)), np.arange(len(y_labels)))
z = np.zeros_like(x)

# 设置每个柱体的高度
height = confusion_matrix.values

# 调整标签的位置
x_positions = np.arange(len(x_labels))
y_positions = np.arange(len(y_labels))

# 绘制3D柱状图
fig = plt.figure(figsize=(12, 8))
ax = fig.add_subplot(111, projection='3d')

# 设置颜色梯度并调整透明度
colors = plt.cm.Reds(height / height.max())
colors[:, :, 3] = 0.5  # 将透明度设为50%

# 绘制长方体并在上面标注百分比
for i in range(len(x_labels)):
    for j in range(len(y_labels)):
        xpos = x_positions[i]
        ypos = y_positions[j]
        zpos = 0
        dx = dy = 0.8
        dz = height[j, i]
        
        ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=colors[j, i], shade=True)
        
        # 仅标注大于等于1%的长方体，字体颜色为高亮，且不受透明度影响
        if dz >= 1:
            ax.text(xpos + 0.4, ypos + 0.4, dz + 7, f'{int(dz)}%', ha='center', va='bottom', color='black', fontsize=10, weight='bold', zorder=25)

# 设置坐标轴标签，并调整它们离坐标轴的距离
ax.set_xlabel('Chosen Expression', labelpad=20)  # 增加labelpad来使标签离坐标轴更远
ax.set_ylabel('Intended Expression', labelpad=20)
ax.set_zlabel('Percentage', labelpad=10)

# 设置坐标轴刻度并调整标签位置
ax.set_xticks(np.arange(len(x_labels)) + 0.4)  # 微调位置使得标签居中
ax.set_xticklabels(x_labels, rotation=60, ha='right', fontsize=10, va='center_baseline', position=(0, -0.1))  # 减少position的值使得刻度标签靠近坐标轴
ax.set_yticks(np.arange(len(y_labels)) - 0.2)
ax.set_yticklabels(y_labels, fontsize=10, va='center_baseline', position=(-0.1, 0))  # 减少position的值使得刻度标签靠近坐标轴

# 设置标题并调整它与图的主要部分的距离
ax.set_title('Percentage of Chosen Emotions per Intended Emotional Expression', pad=15)  # 减少pad来使标题靠近图形

# 调整立方体显示的角度
ax.view_init(elev=30, azim=130)

# 显示图形
plt.show()

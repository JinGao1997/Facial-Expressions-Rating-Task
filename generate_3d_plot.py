import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap
from mpl_toolkits.mplot3d import Axes3D

# 定义每个 Intended Expression 的颜色

expression_colors = {
    'enjoyment': '#D4844A',  # 适度亮的红棕色
    'neutral': '#CC6A48',    # 适度亮的橙棕色
    'disgust': '#DD8F5A',    # 适度亮的橙色
    'affiliation': '#E6A169',# 明亮的橙黄色
    'dominance': '#F0B381',  # 更浅的橙黄色
    'other': '#FFCD9B'       # 更浅的橙色
}


# 读取数据
file_path = 'aligned_data.xlsx'
data = pd.read_excel(file_path)

# 从Material列中提取intended expressions
def get_intended_expression(material):
    if 'enj' in material:
        return 'enjoyment'
    elif 'aff' in material:
        return 'affiliation'
    elif 'dom' in material:
        return 'dominance'
    elif 'dis' in material:
        return 'disgust'
    elif 'neu' in material:
        return 'neutral'
    else:
        return 'other'

data['Intended_Expression'] = data['Material'].apply(get_intended_expression)

# 映射chosen expressions
mapping = {1: 'enjoyment', 2: 'affiliation', 3: 'dominance', 4: 'disgust', 5: 'neutral', 6: 'other'}
data['Chosen_Expression'] = data['Categorizing_Expressions_Score'].map(mapping)

# 生成交叉表，并确保包含所有类别
intended_order = ['neutral', 'enjoyment', 'disgust', 'affiliation', 'dominance']
chosen_order = ['neutral', 'enjoyment', 'disgust', 'affiliation', 'dominance', 'other']

confusion_matrix = pd.crosstab(data['Intended_Expression'], data['Chosen_Expression'], normalize='index') * 100

# 添加缺失的列（例如 'other'），并按指定顺序排序
for col in chosen_order:
    if col not in confusion_matrix.columns:
        confusion_matrix[col] = 0
confusion_matrix = confusion_matrix[chosen_order]

# 对行和列按照指定的顺序进行排序
confusion_matrix = confusion_matrix.reindex(intended_order)

# 准备3D绘图数据
x_labels = confusion_matrix.columns
y_labels = confusion_matrix.index
height = confusion_matrix.values

# 生成 x 和 y 的网格
x, y = np.meshgrid(np.arange(len(x_labels)), np.arange(len(y_labels)))

# 创建 3D 图表
fig = plt.figure(figsize=(12, 8))
ax = fig.add_subplot(111, projection='3d')

# 绘制柱状图，并根据 Intended Expression 设置颜色
for i in range(len(x_labels)):
    for j in range(len(y_labels)):
        xpos = x[j, i]
        ypos = y[j, i]
        zpos = 0
        dx = dy = 0.8
        dz = height[j, i]
        
        # 获取当前 Intended Expression 的颜色
        color = expression_colors[y_labels[j]]
        
        # 绘制柱体
        ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=color, shade=True, zorder=50)

# 绘制百分比标签，并使用黑色
for i in range(len(x_labels)):
    for j in range(len(y_labels)):
        xpos = x[j, i]
        ypos = y[j, i]
        dz = height[j, i]

        if dz >= 5:
            ax.text(xpos + 0.4, ypos + 0.4, dz + 0, f'{int(dz)}%', 
                    ha='center', va='bottom', color='black', fontsize=8.5, weight='bold', zorder=1000)

# 设置轴标签
ax.set_xlabel('Chosen Expression', labelpad=20)
ax.set_ylabel('Intended Expression', labelpad=20)
ax.set_zlabel('Percentage', labelpad=10)

# 设置刻度标签，调整 Intended Expression 标签的位置
ax.set_xticks(np.arange(len(x_labels)) + 0.1)
ax.set_xticklabels(x_labels, rotation=60, ha='right', fontsize=10, va='center_baseline')
ax.set_yticks(np.arange(len(y_labels)) - 0.2)
ax.set_yticklabels(y_labels, fontsize=10, va='center_baseline', position=(-0.15, 0))  # 调整 position 参数

# 设置标题
ax.set_title('Percentage of Chosen Emotions per Intended Emotional Expression', pad=15)

# 调整视角
ax.view_init(elev=32, azim=55)

# 显示图形
plt.show()

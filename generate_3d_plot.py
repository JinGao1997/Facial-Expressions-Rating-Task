import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D

# Optimized color scheme for better discriminability and aesthetics
optimized_colors = {
    'Enjoyment': '#1F77B4',  # Deep Blue
    'Neutral': '#7F7F7F',    # Grey
    'Disgust': '#8C564B',    # Brownish Red
    'Affiliation': '#2CA02C',# Light Green
    'Dominance': '#FF7F0E',  # Deep Orange
    'Other': '#9467BD'       # Purple
}

# Label text color
label_color = '#FFFFFF'  # White

# Read data from the provided file
file_path = 'aligned_data.xlsx'  # Replace with your actual file path
data = pd.read_excel(file_path)

# Extract intended expressions from the Material column
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

# Map chosen expressions from scores to labels
mapping = {1: 'Enjoyment', 2: 'Affiliation', 3: 'Dominance', 4: 'Disgust', 5: 'Neutral', 6: 'Other'}
data['Chosen_Expression'] = data['Categorizing_Expressions_Score'].map(mapping)

# Generate confusion matrix with specified order and normalize by index
intended_order = ['Neutral', 'Enjoyment', 'Disgust', 'Affiliation', 'Dominance']
chosen_order = ['Neutral', 'Enjoyment', 'Disgust', 'Affiliation', 'Dominance', 'Other']
confusion_matrix = pd.crosstab(data['Intended_Expression'], data['Chosen_Expression'], normalize='index') * 100

# Ensure all columns are present and in the correct order
for col in chosen_order:
    if col not in confusion_matrix.columns:
        confusion_matrix[col] = 0
confusion_matrix = confusion_matrix[chosen_order]
confusion_matrix = confusion_matrix.reindex(intended_order)

# Save confusion matrix for reference
confusion_matrix.to_csv('confusion_matrix_HitRate.csv')
confusion_matrix.to_excel('confusion_matrix_HitRate.xlsx')

# Prepare 3D plot data
x_labels = confusion_matrix.columns
y_labels = confusion_matrix.index
height = confusion_matrix.values

x, y = np.meshgrid(np.arange(len(x_labels)), np.arange(len(y_labels)))

# Mark special columns
dominance_other = (y_labels == 'Dominance').nonzero()[0][0], (x_labels == 'Other').nonzero()[0][0]
dominance_dominance = (y_labels == 'Dominance').nonzero()[0][0], (x_labels == 'Dominance').nonzero()[0][0]

# Create 3D bar plot
fig = plt.figure(figsize=(12, 8))
ax = fig.add_subplot(111, projection='3d')

# Adjust bar width for better spacing
dx = dy = 0.6

# First, plot the special bars
j, i = dominance_dominance
xpos, ypos, zpos = x[j, i], y[j, i], 0
dz = height[j, i]
color = optimized_colors[y_labels[j]]
ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=color, shade=True, zorder=1)

j, i = dominance_other
xpos, ypos, zpos = x[j, i], y[j, i], 0
dz = height[j, i]
color = optimized_colors[y_labels[j]]
ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=color, shade=True, zorder=2)

# Plot other bars excluding the special ones
for j in range(len(y_labels)):
    for i in range(len(x_labels)):
        if (j, i) in [dominance_other, dominance_dominance]:
            continue
        xpos, ypos, zpos = x[j, i], y[j, i], 0
        dz = height[j, i]
        color = optimized_colors[y_labels[j]]
        ax.bar3d(xpos, ypos, zpos, dx, dy, dz, color=color, shade=True, zorder=0)

# Add percentage labels
for i in range(len(x_labels)):
    for j in range(len(y_labels)):
        xpos, ypos = x[j, i], y[j, i]
        dz = height[j, i]
        if dz >= 5:  # Display for values >= 5%
            ax.text(xpos + 0.4, ypos + 0.4, dz + 1, f'{dz:.1f}%', 
                    ha='center', va='bottom', color=label_color, fontsize=8, weight='bold', zorder=1000)


# Setting axis labels and ticks
ax.set_xlabel('Chosen Expression', labelpad=20, fontsize=12)
ax.set_ylabel('Intended Expression', labelpad=20, fontsize=12)
ax.set_zlabel('Percentage', labelpad=20, fontsize=12)

ax.set_xticks(np.arange(len(x_labels)) + 0.1)
ax.set_xticklabels(x_labels, rotation=60, ha='right', fontsize=9.5, va='center_baseline')
ax.set_yticks(np.arange(len(y_labels)) - 0.2)
ax.set_yticklabels(y_labels, fontsize=9.5, va='center_baseline')

# Title and view adjustments
ax.set_title('Percentage of Chosen Emotions\nper Intended Emotional Expression', pad=15)
ax.view_init(elev=30, azim=45)

# Save and display the plot
plt.savefig('confusion_matrix_3d_plot_final.png', dpi=300, format='png', bbox_inches='tight')
plt.show()

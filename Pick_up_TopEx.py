import os
import shutil

# 设置源文件夹路径和目标文件夹路径
source_folder = r'C:\Users\neuro-lab\Top_Expressors\Fema_Top'
destination_folder = r'C:\Users\neuro-lab\Top_Expressors\neuFema_Top'

# 指定需要查找的文件类型和精确匹配的文件名关键词
file_extensions = ['.png']  # 需要查找的文件类型
file_name_exact_keywords = [
    'Fema32', 'Fema46', 'Fema64', 'Fema16', 'Fema30', 
    'Fema10', 'Fema66', 'Fema24', 'Fema4', 'Fema12',
    'Fema60', 'Fema82', 'Fema6', 'Fema68', 'Fema86',
    'Fema56', 'Fema38', 'Fema88', 'Fema80', 'Fema28',
    'Fema58', 'Fema20', 'Fema26', 'Fema52', 'Fema50',
    'Fema40', 'Fema22', 'Fema8', 'Fema78', 'Fema14'
]  # 从图片中提取的精确匹配关键词

# 如果目标文件夹不存在，则创建它
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# 遍历源文件夹中的所有文件
for root, dirs, files in os.walk(source_folder):
    for file in files:
        # 检查文件扩展名是否符合要求
        if any(file.lower().endswith(ext) for ext in file_extensions):
            # 检查文件名是否包含 'aff' 并且包含指定关键词
            if 'neu' in file.lower():
                for keyword in file_name_exact_keywords:
                    if keyword in file:
                        # 构建完整的源文件路径和目标文件路径
                        source_file_path = os.path.join(root, file)
                        destination_file_path = os.path.join(destination_folder, file)
                        
                        # 复制文件到目标文件夹
                        shutil.copy2(source_file_path, destination_file_path)
                        print(f'已复制文件: {source_file_path} -> {destination_file_path}')

print('所有符合条件的文件已成功复制。')

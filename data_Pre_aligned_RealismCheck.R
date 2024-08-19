# 加载所需的库
library(readxl)
library(dplyr)
library(writexl)  # 用于将结果保存为 Excel 文件

# 读取 Excel 数据
file_path <- "aligned_data.xlsx"
data <- read_excel(file_path)

# 从 Material 列中提取表情类型
data <- data %>%
  mutate(
    Expression_Type = case_when(
      grepl("dis", Material) ~ "Disgust",
      grepl("enj", Material) ~ "Enjoyment",
      grepl("aff", Material) ~ "Affiliation",
      grepl("dom", Material) ~ "Dominance",
      grepl("neu", Material) ~ "Neutral",
      TRUE ~ "Other"
    )
  )

# 计算每个材料在不同表情类型下的平均 Realism 评分
mean_realism_scores <- data %>%
  group_by(Material, Expression_Type) %>%
  summarise(Average_Realism_Score = mean(Realism_Score, na.rm = TRUE))

# 筛选出平均得分低于 4 分的材料
low_realism_materials <- mean_realism_scores %>%
  filter(Average_Realism_Score < 4)

# 按材料名称汇总，查看哪些材料至少在一个表情类型中得分低于 4 分
low_realism_materials_summary <- low_realism_materials %>%
  group_by(Material) %>%
  summarise(Low_Score_Count = n(), Min_Realism_Score = min(Average_Realism_Score)) %>%
  filter(Low_Score_Count > 0)

# 输出结果
print(low_realism_materials_summary)

# 保存结果到 Excel 文件
write_xlsx(low_realism_materials_summary, "Low_Realism_Scores_Materials.xlsx")

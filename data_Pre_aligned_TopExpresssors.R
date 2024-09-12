# 加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("ggplot2")
install.packages("stringr")
install.packages("writexl")
install.packages("gridExtra")
install.packages("ggpubr")
install.packages("car")

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(writexl)
library(gridExtra)
library(ggpubr)
library(car)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取 Expressor, 性别 和情绪类型信息
data <- data %>%
  mutate(
    Expression_Type = case_when(
      grepl("dis", Material) ~ "Disgust",
      grepl("enj", Material) ~ "Enjoyment",
      grepl("aff", Material) ~ "Affiliation",
      grepl("dom", Material) ~ "Dominance",
      grepl("neu", Material) ~ "Neutral",
      TRUE ~ "Other"
    ),
    Expressor = str_extract(Material, "Fema\\d+|Male\\d+"),
    Gender = case_when(
      str_detect(Material, "Fema") ~ "Female",
      str_detect(Material, "Male") ~ "Male"
    ),
    Expressor_Short = str_replace(Expressor, "Fema", "F") %>% str_replace("Male", "M")
  )

# Step 3: 筛选并排序 Expressors
target_female_expressors <- c("Fema32", "Fema46", "Fema64", "Fema16", "Fema30", 
                              "Fema10", "Fema66", "Fema24", "Fema4", "Fema12", 
                              "Fema60", "Fema82", "Fema6", "Fema68", "Fema86", 
                              "Fema56", "Fema38", "Fema88", "Fema80", "Fema28", 
                              "Fema58", "Fema20", "Fema26", "Fema52", "Fema50", 
                              "Fema40", "Fema22", "Fema8", "Fema78", "Fema14")

target_male_expressors <- c("Male29", "Male37", "Male73", "Male45", "Male35", 
                            "Male5", "Male53", "Male63", "Male47", "Male81", 
                            "Male49", "Male21", "Male79", "Male67", "Male69", 
                            "Male17", "Male15", "Male65", "Male39", "Male31", 
                            "Male51", "Male59", "Male23", "Male85", "Male61", 
                            "Male71", "Male3", "Male19", "Male27", "Male33")

filtered_data <- data %>%
  filter(Expressor %in% c(target_female_expressors, target_male_expressors)) %>%
  mutate(
    Expressor_Short = factor(Expressor_Short, levels = c(str_replace(target_female_expressors, "Fema", "F"), 
                                                         str_replace(target_male_expressors, "Male", "M")))
  ) %>%
  arrange(Expressor_Short)

# Step 4: 计算 Hit Rate, Avg Arousal, Avg Realism
expression_mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)

filtered_data <- filtered_data %>%
  mutate(
    Correct = case_when(
      Categorizing_Expressions_Score == 6 ~ NA_real_,
      expression_mapping[Expression_Type] == Categorizing_Expressions_Score ~ 1,
      TRUE ~ 0
    )
  )

summary_data <- filtered_data %>%
  group_by(Expressor_Short, Gender) %>%
  summarise(
    Hit_Rate = mean(Correct, na.rm = TRUE),
    Avg_Arousal = mean(Arousal_Score, na.rm = TRUE),
    Avg_Realism = mean(Realism_Score, na.rm = TRUE)
  ) %>%
  ungroup()

# Step 5: 计算 UHR (Unbiased Hit Rate)
uhr_data <- filtered_data %>%
  group_by(Expressor_Short, Gender, Expression_Type, Chosen_Expression = Categorizing_Expressions_Score) %>%
  summarise(n = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, n, fill = 0) %>%
  ungroup()

uhr_data <- uhr_data %>%
  rowwise() %>%
  mutate(
    row_sum = sum(c_across(where(is.numeric)), na.rm = TRUE),
    UHR = ifelse(
      row_sum > 0,
      (get(as.character(expression_mapping[Expression_Type])) / row_sum) * 
        (get(as.character(expression_mapping[Expression_Type])) / sum(get(as.character(expression_mapping[Expression_Type])), na.rm = TRUE)),
      0
    )
  )

uhr_summary <- uhr_data %>%
  group_by(Expressor_Short, Gender) %>%
  summarise(Average_UHR = mean(UHR, na.rm = TRUE)) %>%
  ungroup()

# Step 6: 合并所有结果
final_summary <- summary_data %>%
  left_join(uhr_summary, by = c("Expressor_Short", "Gender"))

# 将数据保存为Excel文件
write_xlsx(final_summary, path = "final_summary.xlsx")


# 描述性统计量计算 ----------------------------------------------------------------

# 加载必要的库
library(dplyr)

# 读取Excel文件
# 请确保您已经安装了readxl包：install.packages("readxl")
library(readxl)
data <- read_excel("final_summary.xlsx")

# 计算描述性统计量
descriptive_stats <- data %>%
  summarise(
    Hit_Rate_mean = mean(Hit_Rate, na.rm = TRUE),
    Hit_Rate_sd = sd(Hit_Rate, na.rm = TRUE),
    Hit_Rate_min = min(Hit_Rate, na.rm = TRUE),
    Hit_Rate_max = max(Hit_Rate, na.rm = TRUE),
    
    Avg_Arousal_mean = mean(Avg_Arousal, na.rm = TRUE),
    Avg_Arousal_sd = sd(Avg_Arousal, na.rm = TRUE),
    Avg_Arousal_min = min(Avg_Arousal, na.rm = TRUE),
    Avg_Arousal_max = max(Avg_Arousal, na.rm = TRUE),
    
    Avg_Realism_mean = mean(Avg_Realism, na.rm = TRUE),
    Avg_Realism_sd = sd(Avg_Realism, na.rm = TRUE),
    Avg_Realism_min = min(Avg_Realism, na.rm = TRUE),
    Avg_Realism_max = max(Avg_Realism, na.rm = TRUE),
    
    Average_UHR_mean = mean(Average_UHR, na.rm = TRUE),
    Average_UHR_sd = sd(Average_UHR, na.rm = TRUE),
    Average_UHR_min = min(Average_UHR, na.rm = TRUE),
    Average_UHR_max = max(Average_UHR, na.rm = TRUE)
  )

# 将结果保存为CSV文件
write.csv(descriptive_stats, "Top_Expressors_descriptive_stats.csv", row.names = FALSE)

# 查看结果
print(descriptive_stats)

# 差异检验 --------------------------------------------------------------------
# 标准化数据
data_scaled <- scale(final_summary[, c("Hit_Rate", "Avg_Arousal", "Avg_Realism", "Average_UHR")])

# 计算欧几里得距离矩阵
dist_matrix <- as.matrix(dist(data_scaled))

# 找到具有最大距离的`Expressor_Short`对
max_dist <- max(dist_matrix)
max_dist_indices <- which(dist_matrix == max_dist, arr.ind = TRUE)

# 提取这些行列索引对应的Expressor_Short的实际名称
expr_1 <- final_summary$Expressor_Short[max_dist_indices[1, 1]]
expr_2 <- final_summary$Expressor_Short[max_dist_indices[1, 2]]

# 打印具有最大距离的`Expressor_Short`对
cat("Expressor_Short pair with maximum distance:", as.character(expr_1), "and", as.character(expr_2), "\n")

# 保存最大距离结果
max_dist_result <- data.frame(Expressor_1 = as.character(expr_1), Expressor_2 = as.character(expr_2), Distance = max_dist)
write.csv(max_dist_result, "Max_Distance_TopExpressors.csv", row.names = FALSE)

# 增大画布尺寸以容纳更多的核心热图内容，同时调整标签和字体
png("distance_matrix_heatmap.png", width = 3000, height = 3000, res = 300)
heatmap(dist_matrix, 
        Rowv = NA, Colv = NA, 
        labRow = final_summary$Expressor_Short,  
        labCol = final_summary$Expressor_Short,  
        col = colorRampPalette(c("blue", "white", "red"))(100),  
        main = "Distance Matrix for Top Expressors", 
        xlab = "Expressor", ylab = "Expressor",
        margins = c(4, 4),
        cexRow = 1, cexCol = 1)
dev.off()

# 检验不同性别Expressors的差异 -----------------------------------------------------

# 加载必要的库
install.packages("effsize")
library(readxl)
library(dplyr)
library(effsize)

# 读取Excel文件
data <- read_excel("final_summary.xlsx")

# 分别提取男性和女性的数据
male_data <- data %>% filter(Gender == "Male")
female_data <- data %>% filter(Gender == "Female")

# 定义维度列表
dimensions <- c("Hit_Rate", "Avg_Arousal", "Avg_Realism", "Average_UHR")

# 初始化结果列表
mann_whitney_results <- list()
effect_sizes <- list()

# 进行 Mann-Whitney U 检验和计算 Cohen's d
for (dim in dimensions) {
  # Mann-Whitney U 检验
  test_result <- wilcox.test(male_data[[dim]], female_data[[dim]], exact = FALSE)
  mann_whitney_results[[dim]] <- c("U-statistic" = test_result$statistic, "p-value" = test_result$p.value)
  
  # 计算 Cohen's d
  d_value <- cohen.d(male_data[[dim]], female_data[[dim]])$estimate
  effect_sizes[[dim]] <- d_value
}

# 转换结果为数据框
mann_whitney_df <- do.call(rbind, mann_whitney_results) %>% as.data.frame()
effect_sizes_df <- data.frame(Dimension = names(effect_sizes), `Cohen's d` = unlist(effect_sizes))

# 查看结果
print("Mann-Whitney U Test Results:")
options(scipen = 999)  # 这会减少使用科学计数法的倾向
print(mann_whitney_df)


print("Effect Size (Cohen's d) Results:")
print(effect_sizes_df)

# 保存结果为CSV文件
write.csv(mann_whitney_df, "Mann_Whitney_U_Gender_FourDimensions.csv", row.names = TRUE)
write.csv(effect_sizes_df, "Effect_Size_Cohens_d_Gender_FourDimensions.csv", row.names = FALSE)

## Interpretation of Results:
# - The Mann-Whitney U test results show the U-statistic and p-value for each dimension.
#   A low p-value (typically < 0.05) indicates a statistically significant difference
#   between male and female Expressors in that dimension.
# - Cohen's d provides a measure of the effect size, which can be interpreted as follows:
#   - 0.2: Small effect
#   - 0.5: Medium effect
#   - 0.8: Large effect
#   A positive Cohen's d value suggests that males have higher scores on average,
#   while a negative value indicates higher scores for females.

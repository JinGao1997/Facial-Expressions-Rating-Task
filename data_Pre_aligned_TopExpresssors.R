# 加载必要的库
# 请确保仅在需要时安装包，以避免重复安装
# install.packages(c("readxl", "dplyr", "tidyr", "ggplot2", "stringr", "writexl", "gridExtra", "ggpubr", "car", "pheatmap", "RColorBrewer", "effsize"))

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(writexl)
library(gridExtra)
library(ggpubr)
library(car)
library(pheatmap)
library(RColorBrewer)
library(effsize)
library(tibble)  # 用于处理数据框

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取 Expressor, 性别 和情绪类型信息
data <- data %>%
  mutate(
    Expression_Type = case_when(
      grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
      grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
      grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
      grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
      grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
      TRUE ~ "Other"
    ),
    Expressor = str_extract(Material, "Fema\\d+|Male\\d+"),
    Gender = case_when(
      str_detect(Material, "Fema") ~ "Female",
      str_detect(Material, "Male") ~ "Male",
      TRUE ~ NA_character_
    ),
    Expressor_Short = str_replace_all(Expressor, c("Fema" = "F", "Male" = "M"))
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
    Expressor_Short = factor(Expressor_Short, levels = c(str_replace_all(target_female_expressors, c("Fema" = "F")), 
                                                         str_replace_all(target_male_expressors, c("Male" = "M"))))
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
    ),
    Chosen_Expression = dplyr::recode(Categorizing_Expressions_Score,
                                      `1` = "Enjoyment",
                                      `2` = "Affiliation",
                                      `3` = "Dominance",
                                      `4` = "Disgust",
                                      `5` = "Neutral",
                                      `6` = "Other",
                                      .default = NA_character_)
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
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

uhr_results <- list()

for (expressor in unique(filtered_data$Expressor_Short)) {
  expressor_data <- filtered_data %>% filter(Expressor_Short == expressor)
  
  # 创建混淆矩阵
  confusion_matrix <- expressor_data %>%
    filter(Expression_Type %in% expected_emotions, Chosen_Expression %in% expected_emotions) %>%
    count(Expression_Type, Chosen_Expression) %>%
    spread(key = Chosen_Expression, value = n, fill = 0)
  
  # 补全缺失的行和列
  confusion_matrix <- confusion_matrix %>%
    complete(Expression_Type = expected_emotions, fill = list(n = 0))
  
  confusion_matrix <- confusion_matrix %>%
    column_to_rownames(var = "Expression_Type")
  
  confusion_matrix <- confusion_matrix[expected_emotions, expected_emotions]
  
  N <- sum(confusion_matrix)
  
  uhr_values <- numeric(length(expected_emotions))
  
  for (i in seq_along(expected_emotions)) {
    emotion <- expected_emotions[i]
    a <- confusion_matrix[emotion, emotion]
    b <- sum(confusion_matrix[emotion, ]) - a
    d <- sum(confusion_matrix[, emotion]) - a
    
    if ((a + b) > 0 && (a + d) > 0 && N > 0) {
      UHR <- (a / (a + b)) * (a / (a + d))
    } else {
      UHR <- NA
    }
    
    uhr_values[i] <- UHR
  }
  
  Average_UHR <- mean(uhr_values, na.rm = TRUE)
  
  uhr_results[[expressor]] <- data.frame(
    Expressor_Short = expressor,
    Gender = expressor_data$Gender[1],
    Average_UHR = Average_UHR
  )
}

uhr_summary <- bind_rows(uhr_results)

# Step 6: 合并所有结果
final_summary <- summary_data %>%
  left_join(uhr_summary, by = c("Expressor_Short", "Gender")) %>%
  mutate(
    Expressor_Label = paste0(Expressor_Short, "_", Gender)
  )

# 将数据保存为Excel文件
write_xlsx(final_summary, path = "final_summary.xlsx")

# 描述性统计量计算 ----------------------------------------------------------------

# 计算描述性统计量
descriptive_stats <- final_summary %>%
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

# 检查并处理缺失值
final_summary_clean <- final_summary %>%
  drop_na(Hit_Rate, Avg_Arousal, Avg_Realism, Average_UHR)

# 标准化数据
data_scaled <- scale(final_summary_clean[, c("Hit_Rate", "Avg_Arousal", "Avg_Realism", "Average_UHR")])

# 计算欧几里得距离矩阵
dist_matrix <- as.matrix(dist(data_scaled))

# 设置行名和列名
rownames(dist_matrix) <- final_summary_clean$Expressor_Label
colnames(dist_matrix) <- final_summary_clean$Expressor_Label

# 找到具有最大距离的`Expressor_Label`对
max_dist <- max(dist_matrix)
max_dist_indices <- which(dist_matrix == max_dist, arr.ind = TRUE)

# 提取这些行列索引对应的 Expressor_Label 的实际名称
expr_1 <- rownames(dist_matrix)[max_dist_indices[1, 1]]
expr_2 <- colnames(dist_matrix)[max_dist_indices[1, 2]]

# 打印具有最大距离的 Expressor_Label 对
cat("Expressor pair with maximum distance:", as.character(expr_1), "and", as.character(expr_2), "\n")

# 保存最大距离结果
max_dist_result <- data.frame(Expressor_1 = as.character(expr_1), Expressor_2 = as.character(expr_2), Distance = max_dist)
write.csv(max_dist_result, "Max_Distance_TopExpressors.csv", row.names = FALSE)

# 创建注释数据框，用于在热图中显示性别信息
annotation_df <- final_summary_clean %>%
  select(Expressor_Label, Gender) %>%
  as.data.frame()

rownames(annotation_df) <- annotation_df$Expressor_Label
annotation_df$Expressor_Label <- NULL

# 确认行名匹配
if (!all(rownames(annotation_df) == rownames(dist_matrix))) {
  stop("Row names of annotation_df and dist_matrix do not match.")
}

# 设置颜色方案
heatmap_colors <- colorRampPalette(rev(brewer.pal(n = 11, name = "RdBu")))(100)

# 设置性别的颜色映射
gender_colors <- c(Female = "#E41A1C", Male = "#377EB8")
annotation_colors <- list(Gender = gender_colors)

# 使用 pheatmap 绘制热图
pheatmap(dist_matrix,
         color = heatmap_colors,
         cluster_rows = TRUE,
         cluster_cols = TRUE,
         annotation_row = annotation_df,
         annotation_col = annotation_df,
         annotation_colors = annotation_colors,
         show_rownames = TRUE,
         show_colnames = TRUE,
         fontsize_row = 6,
         fontsize_col = 6,
         legend = TRUE,
         main = "Distance Matrix for Top Expressors",
         fontsize = 10,
         filename = "distance_matrix_heatmap_improved.png",
         width = 10,
         height = 10)

# 检验不同性别Expressors的差异 -----------------------------------------------------

# 分别提取男性和女性的数据
male_data <- final_summary_clean %>% filter(Gender == "Male")
female_data <- final_summary_clean %>% filter(Gender == "Female")

# 定义需要检验的维度列表
dimensions <- c("Hit_Rate", "Avg_Arousal", "Avg_Realism", "Average_UHR")

# 初始化结果列表
mann_whitney_results <- list()
effect_sizes <- list()

# 对每个维度进行检验
for (dim in dimensions) {
  # Mann-Whitney U 检验
  test_result <- wilcox.test(male_data[[dim]], female_data[[dim]], exact = FALSE)
  mann_whitney_results[[dim]] <- c("U-statistic" = test_result$statistic, "p-value" = test_result$p.value)
  
  # 计算 Cohen's d
  d_value <- cohen.d(male_data[[dim]], female_data[[dim]])$estimate
  effect_sizes[[dim]] <- d_value
}

# 将结果转换为数据框
mann_whitney_df <- do.call(rbind, mann_whitney_results) %>% as.data.frame()
effect_sizes_df <- data.frame(Dimension = names(effect_sizes), `Cohen's d` = unlist(effect_sizes))

# 打印结果
print("Mann-Whitney U Test Results:")
options(scipen = 999)  # 避免科学计数法显示
print(mann_whitney_df)

print("Effect Size (Cohen's d) Results:")
print(effect_sizes_df)

# 保存结果为 CSV 文件
write.csv(mann_whitney_df, "Mann_Whitney_U_Gender_FourDimensions.csv", row.names = TRUE)
write.csv(effect_sizes_df, "Effect_Size_Cohens_d_Gender_FourDimensions.csv", row.names = FALSE)

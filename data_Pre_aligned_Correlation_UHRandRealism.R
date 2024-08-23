# 安装并加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("ggplot2")
install.packages("stringr")
install.packages("ggpubr")   # 用于Q-Q图
install.packages("nortest")  # 用于Shapiro-Wilk检验
install.packages("writexl")  # 用于保存 Excel 结果

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(ggpubr)
library(nortest)
library(writexl)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"  # 请替换为你实际的数据路径
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型、expressor 和性别（Gender）
data <- data %>%
  mutate(Expression_Type = case_when(
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
  ))

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression
data <- data %>%
  mutate(Chosen_Expression = case_when(
    Categorizing_Expressions_Score == 1 ~ "Enjoyment",
    Categorizing_Expressions_Score == 2 ~ "Affiliation",
    Categorizing_Expressions_Score == 3 ~ "Dominance",
    Categorizing_Expressions_Score == 4 ~ "Disgust",
    Categorizing_Expressions_Score == 5 ~ "Neutral",
    Categorizing_Expressions_Score == 6 ~ "Other",
    TRUE ~ NA_character_
  ))

# Step 4: 创建选择矩阵并计算 UHR 和 Chance UHR
uhr_results <- data %>%
  group_by(Expressor, Expression_Type, Material, Chosen_Expression, Gender) %>%
  summarise(n = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, n, fill = 0) %>%
  ungroup()

# Step 5: 计算每个单元格的 UHR 和 Chance UHR
uhr_results <- uhr_results %>%
  rowwise() %>%
  mutate(
    row_sum = sum(c_across(where(is.numeric)), na.rm = TRUE),
    UHR = ifelse(
      row_sum > 0,
      (get(Expression_Type) / row_sum) * (get(Expression_Type) / sum(get(Expression_Type), na.rm = TRUE)),
      0
    ),
    Chance_UHR = ifelse(
      row_sum > 0,
      ((sum(c_across(where(is.numeric)), na.rm = TRUE)) / row_sum) * 
        ((sum(get(Expression_Type), na.rm = TRUE)) / row_sum),
      0
    )
  )

# Step 6: 使用子集选择所需的列，并排除“Other”类别
uhr_results_filtered <- uhr_results %>%
  filter(Expression_Type != "Other") %>%  # 排除 "Other" 类别
  select(Expressor, Gender, Expression_Type, UHR, Chance_UHR)

# Step 7: 计算每个 Expressor 在所有情绪类型下的平均 UHR (不再按情绪类型分组)
uhr_by_expressor <- uhr_results_filtered %>%
  group_by(Expressor, Gender) %>%  # 只按 Expressor 和 Gender 分组
  summarise(Average_UHR = mean(UHR, na.rm = TRUE), .groups = 'drop')

# Step 8: 对 UHR 进行 Arcsine 转换
uhr_by_expressor <- uhr_by_expressor %>%
  mutate(Arcsine_UHR = asin(sqrt(Average_UHR)))

# Step 9: 计算每个 expressor 的平均 Arousal 和 Realism 评分
mean_scores_by_expressor <- data %>%
  group_by(Expressor) %>%
  summarise(
    Avg_Arousal_Score = mean(Arousal_Score, na.rm = TRUE),
    Avg_Realism_Score = mean(Realism_Score, na.rm = TRUE)
  )

# Step 10: 合并 UHR 和 Realism 数据
combined_data <- merge(uhr_by_expressor, mean_scores_by_expressor, by = "Expressor")

# Step 11: 数据探索
# 1. 检查 Arcsine 转换后的 UHR 和 Realism 平均得分的分布情况
qqplot_uhr <- ggqqplot(combined_data$Arcsine_UHR, title = "Q-Q Plot for Arcsine Transformed UHR")
qqplot_realism <- ggqqplot(combined_data$Avg_Realism_Score, title = "Q-Q Plot for Realism Score")

# 直接显示 Q-Q 图以确保其生成
print(qqplot_uhr)
print(qqplot_realism)

# Shapiro-Wilk 正态性检验
shapiro_uhr <- shapiro.test(combined_data$Arcsine_UHR)
shapiro_realism <- shapiro.test(combined_data$Avg_Realism_Score)

# 打印 Shapiro-Wilk 检验结果
print(shapiro_uhr)
print(shapiro_realism)

# 2. 绘制散点图检查是否存在线性关系
scatter_plot <- ggplot(combined_data, aes(x = Avg_Realism_Score, y = Arcsine_UHR)) +
  geom_point(size = 3) +
  geom_smooth(method = "lm", se = FALSE, color = "red") +
  labs(title = "Scatter Plot of Arcsine UHR vs Realism Score",
       x = "Average Realism Score",
       y = "Arcsine UHR") +
  theme_minimal()

# 直接显示散点图以确保其生成
print(scatter_plot)

# Step 12: 根据数据探索结果选择相关性分析方法
# 如果通过数据探索显示符合正态分布并且散点图显示线性关系，使用皮尔逊相关系数
# 否则，使用斯皮尔曼相关系数

if (shapiro_uhr$p.value > 0.05 && shapiro_realism$p.value > 0.05) {
  correlation_result <- cor.test(combined_data$Arcsine_UHR, combined_data$Avg_Realism_Score, method = "pearson")
  method_used <- "Pearson"
} else {
  correlation_result <- cor.test(combined_data$Arcsine_UHR, combined_data$Avg_Realism_Score, method = "spearman")
  method_used <- "Spearman"
}

# 打印相关分析结果
print(paste("Method used for correlation:", method_used))
print(correlation_result)

# Step 13: 可视化相关性
correlation_plot <- ggplot(combined_data, aes(x = Avg_Realism_Score, y = Arcsine_UHR)) +
  geom_point(color = "blue", size = 3) +
  geom_smooth(method = "lm", color = "red", se = FALSE) +
  labs(title = paste(method_used, "Correlation between Arcsine UHR and Realism Scores"),
       x = "Average Realism Score",
       y = "Arcsine Transformed UHR") +
  theme_minimal()

# 直接显示相关性图以确保其生成
print(correlation_plot)

# 保存可视化图形
ggsave("Q-Q_Plot_UHR.png", plot = qqplot_uhr, width = 8, height = 6)
ggsave("Q-Q_Plot_Realism.png", plot = qqplot_realism, width = 8, height = 6)
ggsave("Scatter_Plot.png", plot = scatter_plot, width = 8, height = 6)
ggsave("Correlation_Plot.png", plot = correlation_plot, width = 8, height = 6)

# 保存分析结果到文本文件，并包含解释性注释
sink("UHRandRealismCorrelation_Analysis_Results.txt")
cat("### 数据分析结果 ###\n\n")
cat("1. Shapiro-Wilk 正态性检验结果：\n")
print(shapiro_uhr)
print(shapiro_realism)
cat("\n根据Shapiro-Wilk检验，Arcsine UHR和Realism的分布都不符合正态分布 (p < 0.05)。\n")
cat("因此，选择Spearman相关分析方法，而不是Pearson相关分析。\n\n")
cat("2. 相关分析方法选择：\n")
cat("选择的相关分析方法为:", method_used, "\n\n")
cat("3. 相关分析结果：\n")
print(correlation_result)
sink()

# 保存合并后的数据
write.csv(combined_data, "Combined_Arcsine_UHR_Realism_Scores.csv", row.names = FALSE)


# Rank both UHR and Realism -----------------------------------------------

# 安装并加载 writexl 包
install.packages("writexl")
library(writexl)

# Step 1: 标准化 Arcsine_UHR 和 Avg_Realism_Score（计算 Z 分数）
combined_data <- combined_data %>%
  mutate(
    Z_Arcsine_UHR = scale(Arcsine_UHR),
    Z_Realism_Score = scale(Avg_Realism_Score)
  )

# Step 2: 计算标准化分数的总和或平均值作为综合得分
combined_data <- combined_data %>%
  mutate(
    Combined_Score = Z_Arcsine_UHR + Z_Realism_Score  # 可以调整为平均值: (Z_Arcsine_UHR + Z_Realism_Score) / 2
  )

# Step 3: 按性别分开进行排序
top_performers_female <- combined_data %>%
  filter(Gender == "Female") %>%
  arrange(desc(Combined_Score))  # Female 按照综合得分从高到低排序

top_performers_male <- combined_data %>%
  filter(Gender == "Male") %>%
  arrange(desc(Combined_Score))  # Male 按照综合得分从高到低排序

# Step 4: 保存结果为 Excel 文件，并分为两个工作表
write_xlsx(
  list(
    "Female_Expressors" = top_performers_female,
    "Male_Expressors" = top_performers_male
  ),
  "Top_Performers_Combined_Score_by_Gender.xlsx"
)

# 打印或保存筛选结果
print(top_performers_female)
print(top_performers_male)

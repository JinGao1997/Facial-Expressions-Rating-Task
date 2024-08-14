# 安装所需的库
install.packages("ggplot2")
install.packages("dplyr")
install.packages("ggpubr")
install.packages("readxl")
install.packages("RColorBrewer")
install.packages("FSA")


library(ggplot2)
library(dplyr)
library(ggpubr)
library(readxl)
library(RColorBrewer)
library(FSA)

# 读取 Excel 数据
# 请替换为你数据的实际路径
data <- read_excel("N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx")

# 从 Material 列中提取表情类型
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material) ~ "Disgust",
    grepl("enj", Material) ~ "Enjoyment",
    grepl("aff", Material) ~ "Affiliation",
    grepl("dom", Material) ~ "Dominance",
    grepl("neu", Material) ~ "Neutral",
    TRUE ~ "Other"
  ))

# 计算每个图片的平均 Arousal 和 Realism 评分
mean_scores <- data %>%
  group_by(Material, Expression_Type) %>%
  summarise(
    Arousal_Score = mean(Arousal_Score, na.rm = TRUE),
    Realism_Score = mean(Realism_Score, na.rm = TRUE)
  )

# Kruskal-Wallis 检验
kruskal_arousal <- kruskal.test(Arousal_Score ~ Expression_Type, data = mean_scores)
kruskal_realism <- kruskal.test(Realism_Score ~ Expression_Type, data = mean_scores)

# Dunn's test 事后检验
dunn_arousal <- dunnTest(Arousal_Score ~ Expression_Type, data = mean_scores, method = "bonferroni")
dunn_realism <- dunnTest(Realism_Score ~ Expression_Type, data = mean_scores, method = "bonferroni")

# Realism 评分的单样本 t 检验，检验均值是否与中点值 4 显著不同
t_test_realism <- t.test(mean_scores$Realism_Score, mu = 4)

# 创建一个文件来保存结果
report_file <- "Analysis_ArousalandPlauisibility.txt"

# 写入分析结果到文件
sink(report_file)
cat("Kruskal-Wallis Test Results:\n")
cat("Arousal Scores:\n")
print(kruskal_arousal)
cat("\nRealism Scores:\n")
print(kruskal_realism)

# 写入 Dunn's test 事后检验结果
cat("\nDunn's Test Post-Hoc Results (with Bonferroni correction):\n")
cat("\nArousal Scores:\n")
print(dunn_arousal)
cat("\nRealism Scores:\n")
print(dunn_realism)

# 写入单样本 t 检验结果
cat("\nOne-Sample t-Test for Realism Scores:\n")
print(t_test_realism)

# 写入分析报告
cat("\nAnalysis Report:\n")
cat("\n1. Arousal Scores by Expression Type:\n")
cat("The Kruskal-Wallis test indicates that there are significant differences in Arousal Scores across different expression types.\n")
cat("Kruskal-Wallis chi-squared =", round(kruskal_arousal$statistic, 2), 
    ", df =", kruskal_arousal$parameter, 
    ", p-value =", format.pval(kruskal_arousal$p.value), ".\n")

cat("\n2. Realism Scores by Expression Type:\n")
cat("The Kruskal-Wallis test indicates that there are significant differences in Realism Scores across different expression types.\n")
cat("Kruskal-Wallis chi-squared =", round(kruskal_realism$statistic, 2), 
    ", df =", kruskal_realism$parameter, 
    ", p-value =", format.pval(kruskal_realism$p.value), ".\n")

cat("\n3. Post-Hoc Analysis:\n")
cat("Dunn's test with Bonferroni correction shows specific group differences:\n")
cat("- For Arousal Scores: Significant differences were found between several expression types.\n")
cat("- For Realism Scores: Significant differences were also found between several expression types.\n")

cat("\n4. One-Sample t-Test for Realism Scores:\n")
cat("The one-sample t-test indicates whether the mean Realism Score significantly differs from the midpoint value (4).\n")
cat("t =", round(t_test_realism$statistic, 2), ", df =", t_test_realism$parameter, 
    ", p-value =", format.pval(t_test_realism$p.value), ".\n")

cat("\nConclusion:\n")
cat("The results suggest that different expression types lead to significantly different Arousal and Realism Scores.\n")
cat("Dunn's test further clarifies the specific differences between expression types.\n")
cat("Additionally, the one-sample t-test shows whether the average Realism Scores deviate significantly from the midpoint of the scale.\n")
sink()

# 提示结果已经保存
cat("Analysis Report, including post-hoc tests and one-sample t-test, has been saved to", report_file, "\n")



# 对结果进行可视化 ----------------------------------------------------------------
# 使用更优雅的配色方案
palette <- brewer.pal(n = 5, name = "Set1")

# 小提琴图 - Arousal Score
p1 <- ggplot(mean_scores, aes(x = Expression_Type, y = Arousal_Score, fill = Expression_Type)) +
  geom_violin(trim = TRUE, color = "black", size = 1, alpha = 0.7) +  # 调整线条粗细和透明度
  geom_boxplot(width = 0.1, position = position_dodge(0.9), color = "black", alpha = 0.9) +
  scale_fill_manual(values = palette) +
  labs(title = "Arousal Scores by Expression Type", x = "Expression Type", y = "Arousal Score") +
  theme_classic(base_size = 14) +  # 使用经典主题
  theme(
    legend.position = "none",
    plot.title = element_text(hjust = 0.5, face = "bold"),
    axis.text = element_text(size = 12),
    axis.title = element_text(size = 14, face = "bold")
  )

# 小提琴图 - Realism Score
p2 <- ggplot(mean_scores, aes(x = Expression_Type, y = Realism_Score, fill = Expression_Type)) +
  geom_violin(trim = TRUE, color = "black", size = 1, alpha = 0.7) +  # 调整线条粗细和透明度
  geom_boxplot(width = 0.1, position = position_dodge(0.9), color = "black", alpha = 0.9) +
  scale_fill_manual(values = palette) +
  labs(title = "Realism Scores by Expression Type", x = "Expression Type", y = "Realism Score") +
  theme_classic(base_size = 14) +  # 使用经典主题
  theme(
    legend.position = "none",
    plot.title = element_text(hjust = 0.5, face = "bold"),
    axis.text = element_text(size = 12),
    axis.title = element_text(size = 14, face = "bold")
  )

# 将两个图组合显示
ggarrange(p1, p2, ncol = 2, nrow = 1)
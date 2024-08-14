# 安装并加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("ggplot2")
install.packages("car")      # 用于 ANOVA
install.packages("multcomp") # 用于 Tukey HSD

library(readxl)
library(dplyr)
library(ggplot2)
library(car)
library(multcomp)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取表情类型
data <- data %>%
  mutate(Expression_Type = case_when(
    grepl("dis", Material) ~ "Disgust",
    grepl("enj", Material) ~ "Enjoyment",
    grepl("aff", Material) ~ "Affiliation",
    grepl("dom", Material) ~ "Dominance",
    grepl("neu", Material) ~ "Neutral",
    TRUE ~ "Other"
  ))

# Step 3: 将 Expression_Type 映射到 Categorizing_Expressions_Score
mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)
data <- data %>%
  mutate(Correct = ifelse(mapping[Expression_Type] == Categorizing_Expressions_Score, 1, 0))

# Step 4: 计算 Percentage Hit Rate
percentage_hit_rate <- data %>%
  group_by(Material, Expression_Type) %>%
  summarise(Hit_Rate = mean(Correct, na.rm = TRUE)) %>%
  ungroup()

# Step 5: 计算描述性统计信息
hit_rate_summary <- percentage_hit_rate %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_Hit_Rate = mean(Hit_Rate),
    SD_Hit_Rate = sd(Hit_Rate),
    N = n()
  )

# Step 6: 执行 ANOVA 分析
anova_model <- aov(Hit_Rate ~ Expression_Type, data = percentage_hit_rate)
anova_summary <- summary(anova_model)

# Step 7: 执行 Tukey HSD 事后检验
tukey_hsd <- TukeyHSD(anova_model)

# Step 8: 将统计结果保存到文本文件
sink("statistical_results.txt")
cat("Descriptive Statistics for Percentage Hit Rate:\n")
print(hit_rate_summary)
cat("\nANOVA Results for Hit Rate:\n")
print(anova_summary)
cat("\nTukey HSD Post-Hoc Test Results:\n")
print(tukey_hsd)
sink()

# Step 9: 可视化 Percentage Hit Rate 的条形图，优化比例
p <- ggplot(percentage_hit_rate, aes(x = Expression_Type, y = Hit_Rate, fill = Expression_Type)) +
  geom_bar(stat = "identity", color = "black", position = position_dodge(), width = 0.6) +
  geom_errorbar(aes(ymin = Hit_Rate - sd(Hit_Rate), ymax = Hit_Rate + sd(Hit_Rate)), width = 0.2, position = position_dodge(0.9)) +
  scale_y_continuous(labels = scales::percent_format(accuracy = 1), limits = c(0, 1)) +
  scale_fill_brewer(palette = "Set3") +
  labs(title = "Bar Plot of Mean Percentage Hit Rate by Expression Type",
       x = "Expression Type", y = "Mean Percentage Hit Rate") +
  theme_minimal(base_size = 15) +
  theme(legend.position = "none",
        plot.title = element_text(hjust = 0.5, face = "bold"),
        axis.title.x = element_text(face = "bold"),
        axis.title.y = element_text(face = "bold"))

# 打印图形到屏幕
print(p)

# 使用 ggsave 保存图形，调整宽高比
ggsave("percentage_hit_rate_plot.png", plot = p, width = 10, height = 6, dpi = 300)

# 打印完成信息
cat("Analysis complete. Results saved to 'statistical_results.txt' and 'percentage_hit_rate_plot.png'.\n")

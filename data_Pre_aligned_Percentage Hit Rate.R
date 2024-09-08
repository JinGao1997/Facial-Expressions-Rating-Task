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

# Step 5: 对 Hit Rate 进行 Arcsine 变换
percentage_hit_rate <- percentage_hit_rate %>%
  mutate(Arcsine_Hit_Rate = asin(sqrt(Hit_Rate)))

# Step 6: 计算描述性统计信息（基于 Arcsine 变换后的数据）
hit_rate_summary <- percentage_hit_rate %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_Arcsine_Hit_Rate = mean(Arcsine_Hit_Rate),
    SD_Arcsine_Hit_Rate = sd(Arcsine_Hit_Rate),
    N = n()
  )

# 输出描述性统计结果
print(hit_rate_summary)

# Step 7: 执行 ANOVA 分析（基于 Arcsine 变换后的数据）
anova_model <- aov(Arcsine_Hit_Rate ~ Expression_Type, data = percentage_hit_rate)
anova_summary <- summary(anova_model)

# Step 8: 执行 Tukey HSD 事后检验
tukey_hsd <- TukeyHSD(anova_model)

# Step 9: 将统计结果保存到文本文件
sink("Percentage_Hit_Rate_Arcsine_ANOVA.txt")
cat("Descriptive Statistics for Arcsine Transformed Percentage Hit Rate:\n")
print(hit_rate_summary)
cat("\nANOVA Results for Arcsine Transformed Hit Rate:\n")
print(anova_summary)
cat("\nTukey HSD Post-Hoc Test Results:\n")
print(tukey_hsd)
sink()

# Step 10: 可视化 Arcsine Hit Rate 的条形图
p <- ggplot(percentage_hit_rate, aes(x = Expression_Type, y = Arcsine_Hit_Rate, fill = Expression_Type)) +
  geom_bar(stat = "identity", color = "black", position = position_dodge(), width = 0.6) +
  geom_errorbar(aes(ymin = Arcsine_Hit_Rate - sd(Arcsine_Hit_Rate), ymax = Arcsine_Hit_Rate + sd(Arcsine_Hit_Rate)), 
                width = 0.2, position = position_dodge(0.9)) +
  labs(title = "Bar Plot of Arcsine Transformed Hit Rate by Expression Type",
       x = "Expression Type", y = "Arcsine Transformed Hit Rate") +
  theme_minimal(base_size = 15) +
  theme(legend.position = "none",
        plot.title = element_text(hjust = 0.5, face = "bold"),
        axis.title.x = element_text(face = "bold"),
        axis.title.y = element_text(face = "bold"))

# 打印图形到屏幕
print(p)

# 使用 ggsave 保存图形，调整宽高比
ggsave("arcsine_hit_rate_plot.png", plot = p, width = 10, height = 6, dpi = 300)

# 打印完成信息
cat("Analysis complete. Results saved to 'Percentage_Hit_Rate_Arcsine_ANOVA.txt' and 'arcsine_hit_rate_plot.png'.\n")

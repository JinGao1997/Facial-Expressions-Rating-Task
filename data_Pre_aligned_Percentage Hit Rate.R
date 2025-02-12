# ---------------------------------------------------------------------------
# 安装并加载所需的库
#（如果已经安装，可以注释安装部分）
if(!require("readxl")) install.packages("readxl")
if(!require("dplyr")) install.packages("dplyr")
if(!require("ggplot2")) install.packages("ggplot2")
if(!require("car")) install.packages("car")      # 用于 ANOVA
if(!require("multcomp")) install.packages("multcomp") # 用于 Tukey HSD

library(readxl)
library(dplyr)
library(ggplot2)
library(car)
library(multcomp)

# ---------------------------------------------------------------------------
# Step 1: 读取数据 / Step 1: Read the data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# ---------------------------------------------------------------------------
# Step 2: 从 Material 列中提取表情类型 / Extract expression type from the Material column
data <- data %>% 
  mutate(Expression_Type = case_when(
    grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
    grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
    grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
    grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
    grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
    TRUE ~ "Other"
  ))

# ---------------------------------------------------------------------------
# Step 3: 将 Expression_Type 映射到 Categorizing_Expressions_Score 
# (判断受试选择是否正确) / Map Expression_Type to Categorizing_Expressions_Score
mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)
data <- data %>% 
  mutate(Correct = ifelse(mapping[Expression_Type] == Categorizing_Expressions_Score, 1, 0))

# ---------------------------------------------------------------------------
# Step 4: 计算 Percentage Hit Rate / Calculate percentage hit rate
percentage_hit_rate <- data %>% 
  group_by(Material, Expression_Type) %>% 
  summarise(Hit_Rate = mean(Correct, na.rm = TRUE)) %>% 
  ungroup()

# ---------------------------------------------------------------------------
# Step 5: 对 Hit Rate 进行 Arcsine 变换 / Perform Arcsine transformation on the hit rate
percentage_hit_rate <- percentage_hit_rate %>% 
  mutate(Arcsine_Hit_Rate = asin(sqrt(Hit_Rate)))

# ---------------------------------------------------------------------------
# Step 6: 设置 Expression_Type 为因子，并将 "Neutral" 作为参考水平
# 这一步确保后续模型中，各水平对比均是相对于 "Neutral"
percentage_hit_rate <- percentage_hit_rate %>%
  mutate(Expression_Type = factor(Expression_Type, 
                                  levels = c("Neutral", "Affiliation", "Disgust", "Dominance", "Enjoyment")))

# ---------------------------------------------------------------------------
# Step 7: 计算描述性统计信息（基于 Arcsine 变换后的数据） / Calculate descriptive statistics based on arcsine transformed data
hit_rate_summary <- percentage_hit_rate %>% 
  group_by(Expression_Type) %>% 
  summarise(
    Mean_Arcsine_Hit_Rate = mean(Arcsine_Hit_Rate, na.rm = TRUE),
    SD_Arcsine_Hit_Rate = sd(Arcsine_Hit_Rate, na.rm = TRUE),
    N = n()
  )

# 输出描述性统计结果 / Output descriptive statistics
print(hit_rate_summary)

# ---------------------------------------------------------------------------
# Step 8: 执行 ANOVA 分析（基于 Arcsine 变换后的数据） / Perform ANOVA on arcsine transformed data
anova_model <- aov(Arcsine_Hit_Rate ~ Expression_Type, data = percentage_hit_rate)
anova_summary <- summary(anova_model)

# ---------------------------------------------------------------------------
# Step 9: 执行 Tukey HSD 事后检验 / Perform Tukey HSD post-hoc test
tukey_hsd <- TukeyHSD(anova_model)

# ---------------------------------------------------------------------------
# Step 10: 将统计结果保存到文本文件 / Save statistical results to a text file
sink("Percentage_Hit_Rate_Arcsine_ANOVA.txt")
cat("Descriptive Statistics for Arcsine Transformed Percentage Hit Rate:\n")
print(hit_rate_summary)
cat("\nANOVA Results for Arcsine Transformed Hit Rate:\n")
print(anova_summary)
cat("\nTukey HSD Post-Hoc Test Results:\n")
print(tukey_hsd)
sink()

# 打印完成信息 / Print completion message
cat("Analysis complete. Results saved to 'Percentage_Hit_Rate_Arcsine_ANOVA.txt'\n")

# 加载必要的包
library(readxl)
library(dplyr)
library(lme4)      # GLMM 模型
library(emmeans)   # 事后比较
library(DHARMa)    # 模型诊断

# 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# 数据预处理：转换性别、提取 Expression_Type、生成 Correct 指标等
data <- data %>%
  mutate(
    Rater_Gender = case_when(
      Gender == 1 ~ "Female",
      Gender == 2 ~ "Male",
      TRUE ~ NA_character_
    ),
    Expression_Type = case_when(
      grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
      grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
      grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
      grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
      grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
      TRUE ~ "Other"
    ),
    Chosen_Expression = case_when(
      Categorizing_Expressions_Score == 1 ~ "Enjoyment", 
      Categorizing_Expressions_Score == 2 ~ "Affiliation", 
      Categorizing_Expressions_Score == 3 ~ "Dominance", 
      Categorizing_Expressions_Score == 4 ~ "Disgust", 
      Categorizing_Expressions_Score == 5 ~ "Neutral", 
      Categorizing_Expressions_Score == 6 ~ "Other",
      TRUE ~ NA_character_
    ),
    # 生成 Correct 指标：利用映射将 Expression_Type 与 Categorizing_Expressions_Score 对比
    Correct = ifelse(
      c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, 
        "Disgust" = 4, "Neutral" = 5)[Expression_Type] == Categorizing_Expressions_Score,
      1, 0),
    Face_Gender = case_when(
      grepl("fema", Material, ignore.case = TRUE) ~ "Female",  
      grepl("male", Material, ignore.case = TRUE) ~ "Male",
      TRUE ~ NA_character_
    ),
    Version = case_when(
      grepl("L", Material, ignore.case = TRUE) ~ "L",
      grepl("R", Material, ignore.case = TRUE) ~ "R",
      TRUE ~ NA_character_
    )
  )

# 仅保留 Expression_Type 不为 "Other" 的记录
data_glmm <- data %>% filter(Expression_Type != "Other")

# 设置 Expression_Type 为因子，设定参考水平 (例如将 "Neutral" 作为参考)
expected_emotions <- c("Neutral", "Affiliation", "Disgust", "Dominance", "Enjoyment")
data_glmm <- data_glmm %>% 
  mutate(Expression_Type = factor(Expression_Type, levels = expected_emotions))

# 建立 GLMM：使用 Correct (0/1)作为响应变量，
# 固定效应： Expression_Type, Group, Rater_Gender, Face_Gender, Version
# 随机效应：评分者 (CASE)
glmm_model_shr <- glmer(Correct ~ Expression_Type + Group + Rater_Gender + Face_Gender + Version + (1 | CASE),
                        data = data_glmm,
                        family = binomial,
                        control = glmerControl(optimizer = "bobyqa"))
# 输出模型摘要
model_summary <- summary(glmm_model_shr)
print(model_summary)

# 使用 DHARMa 进行模型诊断
simulationOutput <- simulateResiduals(fittedModel = glmm_model_shr, n = 1000)
# 绘制残差分布图
plot(simulationOutput)
# 检查过度离散（overdispersion）
testOverdispersion(simulationOutput)
# 检查零膨胀（zero-inflation）
testZeroInflation(simulationOutput)

# 事后比较：使用 emmeans 对 Expression_Type 进行多重比较
emms_shr <- emmeans(glmm_model_shr, ~ Expression_Type)
pairwise_shr <- pairs(emms_shr)
print(pairwise_shr)

# 保存所有输出到文本文件
sink("Simple_Hit_Rate_GLMM_DHARMa.txt")
cat("GLMM 模型摘要:\n")
print(model_summary)

cat("\nDHARMa 模型诊断结果:\n")
print(simulationOutput)
cat("\nOverdispersion 检验:\n")
print(testOverdispersion(simulationOutput))
cat("\nZero-Inflation 检验:\n")
print(testZeroInflation(simulationOutput))

cat("\n事后多重比较 (Expression_Type):\n")
print(pairwise_shr)
sink()

cat("所有分析结果已保存到 'Simple_Hit_Rate_GLMM_DHARMa.txt' 文件中。\n")

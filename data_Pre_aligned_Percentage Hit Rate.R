# ---------------------------------------------------------------------------
# 安装并加载所需的库（如果尚未安装，则安装；已安装可注释安装部分）
if(!require("readxl")) install.packages("readxl")
if(!require("dplyr")) install.packages("dplyr")
if(!require("tidyr")) install.packages("tidyr")
if(!require("lmerTest")) install.packages("lmerTest")
if(!require("emmeans")) install.packages("emmeans")
if(!require("ggplot2")) install.packages("ggplot2")

library(readxl)
library(dplyr)
library(tidyr)
library(lmerTest)
library(emmeans)
library(ggplot2)

# ---------------------------------------------------------------------------
# Step 1: 读取数据 / Read data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# ---------------------------------------------------------------------------
# Step 1.1: 转换评分者性别 / Convert Rater_Gender from raw Gender column
data <- data %>% 
  mutate(Rater_Gender = case_when(
    Gender == 1 ~ "Female",
    Gender == 2 ~ "Male",
    TRUE ~ NA_character_
  ))

# ---------------------------------------------------------------------------
# 【调试提示】（可选）检查 Material 字段中是否包含 "male" 和 "fema"
# debug_materials <- data %>% 
#   mutate(has_male = grepl("male", Material, ignore.case = TRUE),
#          has_fema = grepl("fema", Material, ignore.case = TRUE)) %>% 
#   select(Material, has_male, has_fema)
# print(head(debug_materials, 20))

# ---------------------------------------------------------------------------
# Step 2: 从 Material 提取表情类别 / Extract Expression_Type from Material
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

# ---------------------------------------------------------------------------
# 新增步骤：生成 Correct 指标
mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)
data <- data %>% 
  mutate(Correct = ifelse(mapping[Expression_Type] == Categorizing_Expressions_Score, 1, 0))

# ---------------------------------------------------------------------------
# Step 3.1: 从 Material 提取其他变量
data <- data %>% 
  mutate(
    # Face_Gender：先匹配 "fema"（例如 R_enjFema18），再匹配 "male"（例如 L_domMale89 或 L_affMale41）
    Face_Gender = case_when(
      grepl("fema", Material, ignore.case = TRUE) ~ "Female",  
      grepl("male", Material, ignore.case = TRUE) ~ "Male",
      TRUE ~ NA_character_
    ),
    # Version：根据 Material 中是否包含 "L" 或 "R"
    Version = case_when(
      grepl("L", Material, ignore.case = TRUE) ~ "L",
      grepl("R", Material, ignore.case = TRUE) ~ "R",
      TRUE ~ NA_character_
    )
    # Group 直接沿用原始数据中的 Group 变量
  )

# ---------------------------------------------------------------------------
# Step 4: 计算 Simple Hit Rate (SHR)
# 只考虑 Expression_Type 不为 "Other" 的记录
hit_rate_data <- data %>% 
  filter(Expression_Type != "Other") %>% 
  group_by(CASE, Expression_Type, Rater_Gender, Face_Gender, Version, Group) %>% 
  summarise(Hit_Rate = mean(Correct, na.rm = TRUE), .groups = "drop")

# ---------------------------------------------------------------------------
# Step 5: 对 Hit Rate 进行 Arcsine 变换
hit_rate_data <- hit_rate_data %>% 
  mutate(Arcsine_Hit_Rate = asin(sqrt(Hit_Rate)))

# ---------------------------------------------------------------------------
# Step 6: 将 Expression_Type 设为因子，并将 "Neutral" 设为参考水平
expected_emotions <- c("Neutral", "Affiliation", "Disgust", "Dominance", "Enjoyment")
hit_rate_data <- hit_rate_data %>%
  mutate(Expression_Type = factor(Expression_Type, levels = expected_emotions))

# ---------------------------------------------------------------------------
# Step 7: 使用混合效应模型 (LMM) 对 SHR 进行分析
# 固定效应：Expression_Type, Group, Rater_Gender, Face_Gender, Version
# 随机效应：CASE（评分者）
lmm_model_shr <- lmer(Arcsine_Hit_Rate ~ Expression_Type + Group + Rater_Gender + Face_Gender + Version + (1|CASE),
                      data = hit_rate_data)
summary(lmm_model_shr)

# ---------------------------------------------------------------------------
# Step 8: 使用 emmeans 对 Expression_Type 进行事后比较
emms_shr <- emmeans(lmm_model_shr, ~ Expression_Type)
pairwise_shr <- pairs(emms_shr)
print(pairwise_shr)

# ---------------------------------------------------------------------------
# Step 9: 将结果保存到文本文件
sink("Simple_Hit_Rate_Arcsine_LMM.txt")
cat("Descriptive Statistics for Arcsine Transformed Simple Hit Rate:\n")
print(hit_rate_data %>% 
        group_by(Expression_Type) %>% 
        summarise(
          Mean = mean(Arcsine_Hit_Rate, na.rm = TRUE),
          SD = sd(Arcsine_Hit_Rate, na.rm = TRUE),
          N = n()
        ))
cat("\nLMM Results for Arcsine Transformed Simple Hit Rate:\n")
print(summary(lmm_model_shr))
cat("\nPost-Hoc Comparisons for Expression_Type:\n")
print(pairwise_shr)
sink()

cat("Analysis complete. Results saved to 'Simple_Hit_Rate_Arcsine_LMM.txt'\n")

# 安装并加载所需的库
install.packages("lme4")
install.packages("lmerTest")
install.packages("readxl")
install.packages("dplyr")
install.packages("irr")  # 用于计算 Fleiss' Kappa
install.packages("tidyr") # 用于数据转换

library(lme4)
library(lmerTest)
library(readxl)
library(dplyr)
library(irr)
library(tidyr)

print("Script is running...")

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
my_data <- read_excel(file_path)

# 检查数据是否成功加载
if (exists("my_data")) {
  print("Data successfully loaded.")
} else {
  stop("Error: Data not found. Please check the file path and ensure the data is correctly loaded.")
}

# Step 2: 根据 Material 列的信息分解为多个维度
my_data <- my_data %>%
  mutate(Version = ifelse(grepl("^L", Material), "L", "R"),
         Gender = ifelse(grepl("Fema", Material), "Female", "Male"),
         Emotion = case_when(
           grepl("dis", Material) ~ "Disgust",
           grepl("enj", Material) ~ "Enjoyment",
           grepl("aff", Material) ~ "Affiliation",
           grepl("dom", Material) ~ "Dominance",
           grepl("neu", Material) ~ "Neutral"
         ))

# Step 3: 初始化结果数据框
results <- data.frame(
  Group = character(),
  Dimension = character(),
  Group_Difference_Significant = logical(),
  ICC_2_k = numeric(),
  k = numeric(),
  Fleiss_Kappa = numeric(),
  stringsAsFactors = FALSE
)

# Step 4: 检验组间差异、计算 ICC(2,k) 和 Fleiss' Kappa
for (dimension in c("Arousal_Score", "Realism_Score", "Categorizing_Expressions_Score")) {
  
  print(paste("Running mixed-effects model for Dimension:", dimension))
  
  if (dimension != "Categorizing_Expressions_Score") {
    # Step 4.1: 使用 lmer 进行混合效应模型并计算 p-value，加入性别和情绪固定效应
    model <- lmer(as.formula(paste(dimension, "~ Group + Gender + Emotion + (1|CASE) + (1|Material)")), data = my_data)
    model_summary <- summary(model)
    
    # 提取 Group 的 p-value
    group_p_value <- anova(model)["Group", "Pr(>F)"]
    
    # 打印 p-value 以检查其值
    print(paste("p-value for Group:", group_p_value))
    
    # 检查 p-value 是否有效并确定是否继续计算 ICC(2,k)
    if (!is.na(group_p_value)) {
      if (group_p_value < 0.05) {
        print(paste("Group differences are significant for", dimension, "(p-value:", group_p_value, "). ICC(2,k) will not be calculated."))
        results <- rbind(results,
                         data.frame(Group = NA,
                                    Dimension = dimension,
                                    Group_Difference_Significant = TRUE,
                                    ICC_2_k = NA,
                                    k = NA,
                                    Fleiss_Kappa = NA))
        next
      } else {
        print(paste("Group differences are not significant for", dimension, "(p-value:", group_p_value, "). Proceeding to calculate ICC(2,k)."))
      }
    } else {
      print("p-value is NA, skipping ICC(2,k) calculation.")
      results <- rbind(results,
                       data.frame(Group = NA,
                                  Dimension = dimension,
                                  Group_Difference_Significant = NA,
                                  ICC_2_k = NA,
                                  k = NA,
                                  Fleiss_Kappa = NA))
      next
    }
  }
  
  # 获取所有Group的名称
  groups <- unique(my_data$Group)
  
  for (group in groups) {
    print(paste("Running model for Group:", group, "Dimension:", dimension))
    
    # Step 4.2: 准备数据
    group_data <- my_data %>%
      filter(Group == group) %>%
      select(CASE, Material, Version, Gender, Emotion, Categorizing_Expressions_Score, all_of(dimension)) %>%
      filter(!is.na(!!sym(dimension)))
    
    # 确保group_data中有足够的行进行模型拟合
    if (nrow(group_data) < 2) {
      print(paste("Skipping Group:", group, "Dimension:", dimension, "due to insufficient data"))
      next
    }
    
    if (dimension != "Categorizing_Expressions_Score") {
      # Step 4.3: 使用 lme4 包构建模型并计算 ICC(2,k)
      lme4_model <- lmer(as.formula(paste(dimension, "~ 1 + (1|CASE) + (1|Material)")), data = group_data)
      var_components <- as.data.frame(VarCorr(lme4_model))
      
      # 计算 ICC(2,k)
      num_raters <- length(unique(group_data$CASE))  # 计算k值
      icc_2_k <- var_components$vcov[1] / (var_components$vcov[1] + (var_components$vcov[2] / num_raters))
      
      # 存储结果
      results <- rbind(results,
                       data.frame(Group = group,
                                  Dimension = dimension,
                                  Group_Difference_Significant = FALSE,
                                  ICC_2_k = icc_2_k,
                                  k = num_raters,
                                  Fleiss_Kappa = NA))
    } else {
      # Step 4.4: 计算Categorizing_Expressions_Score的Fleiss' Kappa
      kappa_data <- group_data %>%
        select(CASE, Material, Categorizing_Expressions_Score) %>%
        filter(!is.na(Categorizing_Expressions_Score))
      
      if (nrow(kappa_data) > 1) {
        wide_kappa_data <- kappa_data %>%
          pivot_wider(names_from = CASE, values_from = Categorizing_Expressions_Score, values_fill = NA_real_)
        
        if (ncol(wide_kappa_data) > 1) {  # 确保至少有两个评分者的有效数据
          kappa_result <- kappam.fleiss(as.matrix(wide_kappa_data[,-1]))
          kappa_value <- kappa_result$value
        } else {
          kappa_value <- NA
        }
      } else {
        kappa_value <- NA
      }
      
      # 存储Fleiss' Kappa结果
      results <- rbind(results,
                       data.frame(Group = group,
                                  Dimension = dimension,
                                  Group_Difference_Significant = FALSE,
                                  ICC_2_k = NA,
                                  k = NA,
                                  Fleiss_Kappa = kappa_value))
    }
  }
}

# Step 5: 计算组内平均值
average_results <- results %>%
  group_by(Dimension) %>%
  summarise(
    Group_Difference_Significant = NA,
    ICC_2_k = mean(ICC_2_k, na.rm = TRUE),
    k = mean(k, na.rm = TRUE),
    Fleiss_Kappa = mean(Fleiss_Kappa, na.rm = TRUE)
  ) %>%
  mutate(Group = "All_Groups_Average")

# 将平均值结果合并回原始结果
results <- bind_rows(results, average_results)

# Step 6: 输出结果
print("Summary of ICC(2,k), Fleiss' Kappa, and Group Difference results:")
print(results)

# Step 7: 将结果保存为CSV文件
write.csv(results, "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/ICC_2_k_and_Fleiss_Kappa_results.csv", row.names = FALSE)

print("Script completed successfully.")

# Interpretation of the parameter ------------------------------------------

#Fleiss' Kappa is a statistical measure used to assess the agreement between multiple raters when classifying multiple items. The Kappa coefficient ranges between -1 and 1, and the interpretation is as follows:

#Interpretation of Kappa Values:
#Kappa < 0: Indicates agreement that is less than what would be expected by chance. This suggests that the raters are not consistent in their ratings, and there may be systematic bias.
#Kappa = 0: Indicates that the agreement between raters is equivalent to what would be expected by chance.
#Kappa > 0: Indicates that the raters' agreement is better than chance, implying some level of consistency.
#Common Interpretation of Kappa Value Ranges (Landis and Koch, 1977):
#0.00 - 0.20: Slight agreement
#0.21 - 0.40: Fair agreement
#0.41 - 0.60: Moderate agreement
#0.61 - 0.80: Substantial agreement
#0.81 - 1.00: Almost perfect agreement

#For Intraclass Correlation Coefficient (ICC), the interpretation depends on the context and the field of study, but general guidelines can be applied. ICC is used to assess the reliability or consistency of measurements made by different raters (or measurements repeated under different conditions).

#Interpretation of ICC Values:
#ICC < 0.5: Indicates poor reliability. The measurements are not consistent across raters or conditions.
#0.5 ≤ ICC < 0.75: Indicates moderate reliability. The measurements are somewhat consistent, but there is still a considerable amount of variability.
#0.75 ≤ ICC < 0.9: Indicates good reliability. The measurements are fairly consistent across raters or conditions.
#ICC ≥ 0.9: Indicates excellent reliability. The measurements are highly consistent across raters or conditions.


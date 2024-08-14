# 加载所需的库
library(dplyr)
library(tidyr)
library(irr)
library(psych)
library(readxl)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
my_data <- read_excel(file_path)

# Step 2: 初始化结果数据框
results <- data.frame(
  Group = integer(),
  Dimension = character(),
  Statistic = character(),
  Value = numeric(),
  stringsAsFactors = FALSE
)

# Step 3: 获取所有Group的名称
groups <- unique(my_data$Group)

# Step 4: 遍历每个Group并计算各个维度的ICC和Kappa
for (group in groups) {
  print(paste("Processing Group:", group))  # 调试信息
  
  # 筛选出当前Group的数据
  group_data <- my_data %>% filter(Group == group)
  
  # 计算Arousal和Realism的ICC
  for (dimension in c("Arousal_Score", "Realism_Score")) {
    
    # 筛选出当前维度的非NA数据
    dimension_data <- group_data %>%
      select(CASE, Material, all_of(dimension)) %>%
      filter(!is.na(!!sym(dimension)))
    
    if (nrow(dimension_data) > 1) {
      # 将数据转换为宽格式（CASE为列）
      wide_data <- dimension_data %>%
        pivot_wider(names_from = CASE, values_from = all_of(dimension), values_fill = NA_real_)
      
      # 检查每列的变异性
      variability <- sapply(wide_data, function(x) length(unique(x)) > 1)
      
      if (all(variability)) {
        # 计算ICC(1,1) 使用irr包
        icc_result <- icc(as.matrix(wide_data[,-1]), model = "oneway", type = "consistency")
        icc_value <- icc_result$value
      } else {
        icc_value <- NA
      }
      
      # 存储结果
      results <- rbind(results,
                       data.frame(Group = group, Dimension = dimension, 
                                  Statistic = "ICC(1,1)", Value = icc_value))
    } else {
      # 如果CASE数量不足，存储NA
      results <- rbind(results,
                       data.frame(Group = group, Dimension = dimension, 
                                  Statistic = "ICC(1,1)", Value = NA))
    }
  }
  
  # 计算Categorizing_Expressions_Score的Fleiss' Kappa
  kappa_data <- group_data %>%
    select(CASE, Material, Categorizing_Expressions_Score) %>%
    filter(!is.na(Categorizing_Expressions_Score))
  
  if (nrow(kappa_data) > 1) {
    # 转换为矩阵形式（CASE为列）
    wide_kappa_data <- kappa_data %>%
      pivot_wider(names_from = CASE, values_from = Categorizing_Expressions_Score, values_fill = NA_real_)
    
    # 计算Fleiss' Kappa
    kappa_result <- kappam.fleiss(as.matrix(wide_kappa_data[,-1]))
    kappa_value <- kappa_result$value
  } else {
    kappa_value <- NA
  }
  
  # 存储结果
  results <- rbind(results,
                   data.frame(Group = group, Dimension = "Categorizing_Expressions_Score", 
                              Statistic = "Fleiss' Kappa", Value = kappa_value))
}

# Step 5: 输出结果
print("Summary of ICC and Kappa results:")
print(results)

# Step 6: 将结果保存为CSV文件
write.csv(results, "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/ICC_and_Kappa_results_by_Group.csv", row.names = FALSE)



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

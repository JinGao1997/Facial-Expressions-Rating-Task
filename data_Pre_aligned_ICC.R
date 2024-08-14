# Step 1: 加载所需的库
library(dplyr)
library(psych)
library(readxl)

# Step 2: 载入数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
my_data <- read_excel(file_path)

# Step 3: 初始化结果数据框
icc_results <- data.frame(
  Item = character(),
  Dimension = character(),
  ICC = numeric(),
  stringsAsFactors = FALSE
)

# Step 4: 获取所有 Material 的名称
items <- unique(my_data$Material)

# Step 5: 遍历每个 Material 并计算各个维度的 ICC
for (item in items) {
  print(paste("Processing item:", item))  # 调试信息
  
  # 筛选出当前 Material 的数据
  item_data <- my_data %>% filter(Material == item)
  
  # 对每个维度分别计算 ICC
  for (dimension in c("Arousal_Score", "Realism_Score")) {
    
    # 筛选出当前维度的非 NA 数据
    dimension_data <- item_data %>%
      select(CASE, all_of(dimension)) %>%
      filter(!is.na(!!sym(dimension)))
    
    # 如果有足够的 CASE 来计算 ICC
    if (nrow(dimension_data) > 1) {
      
      # 将数据转换为矩阵形式（CASE 为列）
      wide_data <- dimension_data %>%
        pivot_wider(names_from = CASE, values_from = all_of(dimension), values_fill = NA_real_)
      
      # 检查每列的变异性
      variability <- sapply(wide_data, function(x) length(unique(x)) > 1)
      print("Variability check (should be TRUE for all columns to calculate ICC):")
      print(variability)
      
      # 计算 ICC，使用 twoway 模型和 consistency 类型
      if (all(variability)) {
        icc_result <- ICC(as.matrix(wide_data), model = "twoway", type = "consistency")
        icc_value <- icc_result$results[1, "ICC2"]  # 注意这里选择 ICC2 类型
      } else {
        icc_value <- NA
      }
      
      # 存储结果
      icc_results <- rbind(icc_results,
                           data.frame(Item = item, Dimension = dimension, ICC = icc_value))
      
    } else {
      # 如果 CASE 数量不足，存储 NA
      icc_results <- rbind(icc_results,
                           data.frame(Item = item, Dimension = dimension, ICC = NA))
    }
  }
  
  print(paste("Finished processing item:", item))  # 调试信息
}

# Step 6: 输出结果
print("Summary of ICC results:")
print(icc_results)

# Step 7: 将结果保存为 CSV 文件
write.csv(icc_results, "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/ICC_results.csv", row.names = FALSE)

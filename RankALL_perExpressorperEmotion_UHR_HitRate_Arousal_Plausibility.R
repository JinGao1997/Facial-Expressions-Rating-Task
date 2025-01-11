# 加载必要的R包 / Load necessary packages
library(readxl)
library(dplyr)
library(tidyr)
library(stringr)

# Step 1: 读取数据 / Read data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 提取表情类型、Expressor 和性别 / Extract Expression_Type, Expressor, and Gender
data <- data %>%
  mutate(
    Expression_Type = case_when(
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
    )
  )

# Step 3: 生成 Chosen_Expression / Generate Chosen_Expression
mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)
data <- data %>%
  mutate(
    Chosen_Expression = case_when(
      Categorizing_Expressions_Score == 1 ~ "Enjoyment",
      Categorizing_Expressions_Score == 2 ~ "Affiliation",
      Categorizing_Expressions_Score == 3 ~ "Dominance",
      Categorizing_Expressions_Score == 4 ~ "Disgust",
      Categorizing_Expressions_Score == 5 ~ "Neutral",
      Categorizing_Expressions_Score == 6 ~ "Other",
      TRUE ~ NA_character_
    )
  )

# 初始化存储结果的数据框 / Initialize a dataframe to store results
results_all <- data.frame(
  Expressor = character(),
  Expression_Type = character(),
  UHR = numeric(),
  Hit_Rate = numeric(),
  Arousal_Score = numeric(),
  Plausibility_Score = numeric(),
  stringsAsFactors = FALSE
)

# Step 4: 定义 UHR 和 Hit Rate 计算函数 / Define UHR and Hit Rate calculation functions
calculate_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]
    b <- sum(conf_matrix[emotion, ], na.rm = TRUE) - a
    d <- sum(conf_matrix[, emotion], na.rm = TRUE) - a
    
    if (is.na(a) || is.na(b) || is.na(d)) {
      UHR <- NA
    } else if ((a + b) > 0 & (a + d) > 0) {
      UHR <- (a / (a + b)) * (a / (a + d))
    } else {
      UHR <- 0
    }
    
    uhr_results[i] <- UHR
  }
  
  return(uhr_results)
}

calculate_hit_rate <- function(conf_matrix) {
  hit_rate_results <- numeric(length(rownames(conf_matrix)))
  
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    total_row <- sum(conf_matrix[emotion, ], na.rm = TRUE)  # 该行的总和
    correct_row <- conf_matrix[emotion, emotion]           # 对角元素
    
    if (total_row > 0) {
      hit_rate_results[i] <- correct_row / total_row
    } else {
      hit_rate_results[i] <- NA  # 如果总行数为0，返回 NA
    }
  }
  
  return(hit_rate_results)
}

# Step 5: 逐个 Expressor 计算指标 / Calculate metrics for each Expressor
expressors <- unique(data$Expressor)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

for (expr in expressors) {
  # 筛选当前 Expressor 的数据 / Filter data for the current Expressor
  expr_data <- data %>% filter(Expressor == expr)
  
  # 创建混淆矩阵 / Create confusion matrix
  conf_matrix <- expr_data %>%
    count(Expression_Type, Chosen_Expression) %>%
    tidyr::spread(key = Chosen_Expression, value = n, fill = 0)
  
  # 补全行 / Complete rows
  conf_matrix <- tidyr::complete(conf_matrix, Expression_Type = expected_emotions, fill = list(n = 0))
  
  # 补全列 / Complete columns
  missing_cols <- setdiff(expected_emotions, colnames(conf_matrix))
  for (col in missing_cols) {
    conf_matrix[[col]] <- 0
  }
  
  # 转换为矩阵并替换 NA 为 0 / Convert to matrix and replace NA with 0
  conf_matrix <- as.data.frame(conf_matrix)
  rownames(conf_matrix) <- conf_matrix$Expression_Type
  conf_matrix <- conf_matrix %>% select(-Expression_Type)
  conf_matrix[is.na(conf_matrix)] <- 0
  conf_matrix <- as.matrix(conf_matrix)
  
  # 计算 UHR 和 Hit Rate / Calculate UHR and Hit Rate
  uhr_results <- calculate_uhr(conf_matrix)
  hit_rate_results <- calculate_hit_rate(conf_matrix)
  
  # 收集结果 / Collect results
  for (i in seq_along(expected_emotions)) {
    results_all <- rbind(results_all, data.frame(
      Expressor = expr,
      Expression_Type = expected_emotions[i],
      UHR = uhr_results[i],
      Hit_Rate = hit_rate_results[i],
      Arousal_Score = mean(expr_data$Arousal_Score[expr_data$Expression_Type == expected_emotions[i]], na.rm = TRUE),
      Plausibility_Score = mean(expr_data$Realism_Score[expr_data$Expression_Type == expected_emotions[i]], na.rm = TRUE)
    ))
  }
}

# Step 6: 添加 Gender 信息并保存结果 / Add Gender information and save results
expressor_gender <- data %>%
  select(Expressor, Gender) %>%
  distinct()

results_all <- results_all %>%
  left_join(expressor_gender, by = "Expressor")

write.csv(results_all, "RankALL_perExpressorperEmotion_UHR_HitRate_Arousal_Plausibility.csv", row.names = FALSE)

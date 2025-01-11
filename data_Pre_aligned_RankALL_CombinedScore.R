# 加载必要的R包
if (!requireNamespace("readxl")) install.packages("readxl")
if (!requireNamespace("dplyr")) install.packages("dplyr")
if (!requireNamespace("tidyr")) install.packages("tidyr")
if (!requireNamespace("ggplot2")) install.packages("ggplot2")
if (!requireNamespace("stringr")) install.packages("stringr")
if (!requireNamespace("writexl")) install.packages("writexl")

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(writexl)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"  # 替换为实际路径
data <- read_excel(file_path)

# Step 2: 提取表情类型、Expressor 和 Expressor 的性别
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
    Expressor = str_extract(Material, "Fema\\d+|Male\\d+"),  # 提取 expressors ID
    Expressor_Gender = case_when(  # 从 Material 提取 expressors 的性别
      str_detect(Material, "Fema") ~ "Female",
      str_detect(Material, "Male") ~ "Male"
    )
  )

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression
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

# 初始化存储结果的数据框
uhr_results_all <- data.frame(
  Expressor = character(),
  Expression_Type = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  Expressor_Gender = character(),
  stringsAsFactors = FALSE
)

# Step 4: 定义 UHR 和 Chance UHR 的计算函数
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  chance_uhr_results <- numeric(length(rownames(conf_matrix)))
  N <- sum(conf_matrix)  # 总样本数
  
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]  # 对角线元素
    b <- sum(conf_matrix[emotion, ]) - a  # 当前行的其他元素
    d <- sum(conf_matrix[, emotion]) - a  # 当前列的其他元素
    
    if ((a + b) > 0 & (a + d) > 0 & N > 0) {
      UHR <- (a / (a + b)) * (a / (a + d))
      Pc <- ((a + b) / N) * ((a + d) / N)
    } else {
      UHR <- 0
      Pc <- 0
    }
    
    uhr_results[i] <- UHR
    chance_uhr_results[i] <- Pc
  }
  
  return(list(UHR = uhr_results, Chance_UHR = chance_uhr_results))
}

# Step 5: 创建用于保存矩阵的文件夹
output_folder <- "Confusion_Matrices"
if (!dir.exists(output_folder)) dir.create(output_folder)

# Step 6: 计算每个 Expressor 的 UHR 和保存混淆矩阵
expressors <- unique(data$Expressor)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")

for (expr in expressors) {
  expr_data <- data %>% filter(Expressor == expr)
  
  conf_matrix <- expr_data %>%
    filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
    count(Expression_Type, Chosen_Expression) %>%
    tidyr::spread(key = Chosen_Expression, value = n, fill = 0) %>%
    tidyr::complete(Expression_Type = expected_emotions, fill = list(n = 0))
  
  conf_matrix_matrix <- as.matrix(conf_matrix[, -1])  # 删除 Expression_Type 列
  rownames(conf_matrix_matrix) <- conf_matrix$Expression_Type
  
  # 保存混淆矩阵
  write.csv(conf_matrix_matrix, file = file.path(output_folder, paste0(expr, "_Confusion_Matrix.csv")), row.names = TRUE)
  
  # 计算 UHR 和 Chance UHR
  uhr_chance_results <- calculate_uhr_chance_uhr(conf_matrix_matrix)
  
  for (i in seq_along(expected_emotions)) {
    uhr_results_all <- rbind(uhr_results_all, data.frame(
      Expressor = expr,
      Expression_Type = expected_emotions[i],
      UHR = uhr_chance_results$UHR[i],
      Chance_UHR = uhr_chance_results$Chance_UHR[i],
      Expressor_Gender = first(expr_data$Expressor_Gender)
    ))
  }
}

# Step 7: 计算 Performance_Above_Chance 和 arcsine 转换
uhr_results_all <- uhr_results_all %>%
  mutate(
    Performance_Above_Chance = UHR - Chance_UHR
  )

# 保存每个 Expressor 的 UHR 详细数据
write.csv(uhr_results_all, "Detailed_UHR_Per_Expressor.csv", row.names = FALSE)

# Step 8: 提取 Plausibility Scores
plausibility_col <- ifelse("Realism_Score" %in% colnames(data), "Realism_Score", "Plausibility_Score")
plausibility_scores <- data %>%
  group_by(Expressor) %>%
  summarise(Plausibility = mean(.data[[plausibility_col]], na.rm = TRUE), .groups = 'drop')

# Step 9: 计算综合得分
average_uhr_per_expressor <- uhr_results_all %>%
  group_by(Expressor, Expressor_Gender) %>%
  summarise(
    Average_UHR = mean(UHR, na.rm = TRUE),
    .groups = 'drop'
  ) %>%
  mutate(Arcsine_UHR = asin(sqrt(Average_UHR)))

combined_data <- average_uhr_per_expressor %>%
  left_join(plausibility_scores, by = "Expressor") %>%
  mutate(
    Z_Arcsine_UHR = scale(Arcsine_UHR),
    Z_Plausibility = scale(Plausibility),
    Combined_Score = (Z_Arcsine_UHR + Z_Plausibility) / 2
  )

# 保存综合得分
write.csv(combined_data, "Final_Combined_Score_Per_Expressor.csv", row.names = FALSE)

# Step 10: 可视化综合得分排名
# 按性别分面可视化
facet_plot <- ggplot(combined_data, aes(x = reorder(Expressor, Combined_Score), y = Combined_Score, fill = Expressor_Gender)) +
  geom_bar(stat = "identity") +
  coord_flip() +
  facet_wrap(~Expressor_Gender, scales = "free_y") +
  labs(title = "Expressor Rankings by Combined Score (Faceted by Gender)",
       x = "Expressor",
       y = "Combined Score",
       fill = "Gender") +
  theme_minimal()

# 保存分面图表
ggsave("Combined_Score_Rankings_Faceted.png", plot = facet_plot, width = 12, height = 8)

cat("生成了按性别分面的排名图表。\n")

# 安装并加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")   # 确保安装和加载 tidyr
install.packages("writexl")

# 加载所需的库
library(readxl)
library(dplyr)
library(tidyr)   # 加载 tidyr
library(writexl)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从Material列中提取预期的情绪类型
data <- data %>%
  mutate(Intended_Expression = case_when(
    grepl("aff", Material) ~ "Affiliation",
    grepl("enj", Material) ~ "Enjoyment",
    grepl("dis", Material) ~ "Disgust",
    grepl("neu", Material) ~ "Neutral",
    grepl("dom", Material) ~ "Dominance",
    TRUE ~ "Other"
  ))

# Step 3: 将 Categorizing_Expressions_Score 映射为 Chosen_Expression
data <- data %>%
  mutate(Chosen_Expression = case_when(
    Categorizing_Expressions_Score == 1 ~ "Enjoyment",
    Categorizing_Expressions_Score == 2 ~ "Affiliation",
    Categorizing_Expressions_Score == 3 ~ "Dominance",
    Categorizing_Expressions_Score == 4 ~ "Disgust",
    Categorizing_Expressions_Score == 5 ~ "Neutral",
    Categorizing_Expressions_Score == 6 ~ "Other",
    TRUE ~ NA_character_  # 如果遇到未匹配的值，将其设为 NA
  ))

# Step 4: 计算每个CASE的反应准确率，包括不同情绪类别的准确率
accuracy_data <- data %>%
  mutate(Correct = ifelse(Chosen_Expression == Intended_Expression & Chosen_Expression != "Other", 1, 0)) %>%
  group_by(CASE) %>%
  summarise(
    Total_Trials = n(),
    Correct_Trials = sum(Correct),
    Accuracy = Correct_Trials / Total_Trials * 100,
    Affiliation_Accuracy = sum(Correct[Intended_Expression == "Affiliation"]) / sum(Intended_Expression == "Affiliation") * 100,
    Enjoyment_Accuracy = sum(Correct[Intended_Expression == "Enjoyment"]) / sum(Intended_Expression == "Enjoyment") * 100,
    Disgust_Accuracy = sum(Correct[Intended_Expression == "Disgust"]) / sum(Intended_Expression == "Disgust") * 100,
    Neutral_Accuracy = sum(Correct[Intended_Expression == "Neutral"]) / sum(Intended_Expression == "Neutral") * 100,
    Dominance_Accuracy = sum(Correct[Intended_Expression == "Dominance"]) / sum(Intended_Expression == "Dominance") * 100
  )

# 打印每个CASE的反应准确率
print(accuracy_data)

# 保存反应准确率为Excel文件
write_xlsx(accuracy_data, path = "CASE_Accuracy_with_Categories.xlsx")

# Step 5: 进一步按材料的情绪类型分类整理
categorize_materials <- function(materials) {
  categorized <- list(
    Affiliation = paste(materials[grepl("aff", materials)], collapse = ", "),
    Enjoyment = paste(materials[grepl("enj", materials)], collapse = ", "),
    Disgust = paste(materials[grepl("dis", materials)], collapse = ", "),
    Neutral = paste(materials[grepl("neu", materials)], collapse = ", "),
    Dominance = paste(materials[grepl("dom", materials)], collapse = ", ")
  )
  return(categorized)
}

# Step 6: 遍历不同的 target_score 值并生成结果
for (target_score in 1:6) {
  
  # 筛选出符合条件的数据并按CASE进行分类整理
  filtered_data <- data %>%
    filter(Categorizing_Expressions_Score == target_score) %>%
    group_by(CASE) %>%
    summarise(Materials = list(Material)) %>%
    mutate(Categorized_Materials = lapply(Materials, categorize_materials)) %>%
    unnest_wider(Categorized_Materials) %>%
    select(CASE, Affiliation, Enjoyment, Disgust, Neutral, Dominance) %>%
    arrange(CASE)
  
  # 打印整理后的数据
  print(paste("查看特定Categorizing_Expressions_Score值", target_score, "时的Materials和CASE:"))
  print(filtered_data)
  
  # Step 7: 保存结果为Excel文件
  output_file <- paste0("Response_Check_for", target_score, ".xlsx")
  write_xlsx(filtered_data, path = output_file)
  
  print(paste("数据已保存为", output_file))
}

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

# Step 2: 进一步按材料的情绪类型分类整理
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

# Step 3: 遍历不同的 target_score 值并生成结果
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
  
  # Step 4: 保存结果为Excel文件
  output_file <- paste0("Response_Check_for", target_score, ".xlsx")
  write_xlsx(filtered_data, path = output_file)
  
  print(paste("数据已保存为", output_file))
}

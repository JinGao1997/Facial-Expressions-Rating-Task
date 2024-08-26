# 加载所需的库
install.packages("readxl")
install.packages("dplyr")
install.packages("tidyr")
install.packages("ggplot2")
install.packages("stringr")
install.packages("writexl")
install.packages("gridExtra")

library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(writexl)
library(gridExtra)

# Step 1: 读取数据
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# Step 2: 从 Material 列中提取 Expressor, 性别 和情绪类型信息
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
    ),
    Expressor_Short = str_replace(Expressor, "Fema", "F") %>% str_replace("Male", "M")  # 使用缩写
  )

# Step 3: 筛选出你关注的Expressors，并按照你提供的顺序排列
target_female_expressors <- c("Fema32", "Fema46", "Fema64", "Fema16", "Fema30", 
                              "Fema10", "Fema66", "Fema24", "Fema4", "Fema12", 
                              "Fema60", "Fema82", "Fema6", "Fema68", "Fema86", 
                              "Fema56", "Fema38", "Fema88", "Fema80", "Fema28", 
                              "Fema58", "Fema20", "Fema26", "Fema52", "Fema50", 
                              "Fema40", "Fema22", "Fema8", "Fema78", "Fema14")

target_male_expressors <- c("Male29", "Male37", "Male73", "Male45", "Male35", 
                            "Male5", "Male53", "Male63", "Male47", "Male81", 
                            "Male49", "Male21", "Male79", "Male67", "Male69", 
                            "Male17", "Male15", "Male65", "Male39", "Male31", 
                            "Male51", "Male59", "Male23", "Male85", "Male61", 
                            "Male71", "Male3", "Male19", "Male27", "Male33")

filtered_data <- data %>%
  filter(Expressor %in% c(target_female_expressors, target_male_expressors)) %>%
  mutate(
    Expressor_Short = factor(Expressor_Short, levels = c(str_replace(target_female_expressors, "Fema", "F"), 
                                                         str_replace(target_male_expressors, "Male", "M")))
  ) %>%
  arrange(Expressor_Short)

# Step 4: 计算 Percentage Hit Rate, Arousal 和 Realism 平均得分
expression_mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)

filtered_data <- filtered_data %>%
  mutate(
    # 标记无法识别的情况
    Correct = case_when(
      Categorizing_Expressions_Score == 6 ~ NA_real_,  # 将无法识别的情况标记为 NA
      expression_mapping[Expression_Type] == Categorizing_Expressions_Score ~ 1,
      TRUE ~ 0
    )
  )

summary_data <- filtered_data %>%
  group_by(Expressor_Short, Gender) %>%
  summarise(
    Hit_Rate = mean(Correct, na.rm = TRUE),  # 计算时排除无法识别的情况
    Avg_Arousal = mean(Arousal_Score, na.rm = TRUE),
    Avg_Realism = mean(Realism_Score, na.rm = TRUE)
  ) %>%
  ungroup()

# Step 5: 计算 UHR (Unbiased Hit Rate)
uhr_data <- filtered_data %>%
  group_by(Expressor_Short, Gender, Expression_Type, Chosen_Expression = Categorizing_Expressions_Score) %>%
  summarise(n = n(), .groups = 'drop') %>%
  spread(Chosen_Expression, n, fill = 0) %>%
  ungroup()

# 计算 UHR
uhr_data <- uhr_data %>%
  rowwise() %>%
  mutate(
    row_sum = sum(c_across(where(is.numeric)), na.rm = TRUE),
    UHR = ifelse(
      row_sum > 0,
      (get(as.character(expression_mapping[Expression_Type])) / row_sum) * 
        (get(as.character(expression_mapping[Expression_Type])) / sum(get(as.character(expression_mapping[Expression_Type])), na.rm = TRUE)),
      0
    )
  )

# 计算每个 Expressor 的平均 UHR
uhr_summary <- uhr_data %>%
  group_by(Expressor_Short, Gender) %>%
  summarise(Average_UHR = mean(UHR, na.rm = TRUE)) %>%
  ungroup()

# Step 6: 合并所有结果
final_summary <- summary_data %>%
  left_join(uhr_summary, by = c("Expressor_Short", "Gender"))

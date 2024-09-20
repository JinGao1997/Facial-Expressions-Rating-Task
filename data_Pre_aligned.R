# 载入必要的库 / Load necessary libraries
install.packages("openxlsx")
library(dplyr)
library(openxlsx)  # 用于保存Excel文件 / Used for saving Excel files
library(readxl)    # 用于读取Excel文件 / Used for reading Excel files

# 读取数据 / Read the data
data <- read_excel("N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/data_JGFacialExpressionsRating_2024-08-14_13-54.xlsx")

# 提取相关列的名称 / Extract relevant column names
case_column <- "CASE"
lg02_column <- "LG02_01"  # 实验组别列 / Experimental group column
sd01_column <- "SD01"     # 性别列 / Gender column
sd02_column <- "SD02_01"  # 年龄列 / Age column
sd20_column <- "SD20"     # 利手列 / Handedness column
sd21_column <- "SD21_01"  # 专业领域列 / Field of study column

# 查找与各变量对应的列名 / Find column names matching the patterns
lg04_columns <- grep("^LG04", names(data), value = TRUE)
ci_columns <- grep("^CI", names(data), value = TRUE)
dc_columns <- grep("^DC", names(data), value = TRUE)
ra_columns <- grep("^RA", names(data), value = TRUE)

# 初始化一个空的数据框来存储对齐后的数据 / Initialize an empty data frame to store the aligned data
aligned_data <- data.frame()

# 对每一组 LG04 对应的 CI, DC, RA 列进行处理 / Process each group of LG04, CI, DC, and RA columns
for (i in seq_along(lg04_columns)) {
  temp_df <- data.frame(
    CASE = data[[case_column]],                 # 被试的编号 / Participant ID
    Group = data[[lg02_column]],                # 实验组别 / Experimental group
    Gender = data[[sd01_column]],               # 性别 / Gender
    Age = data[[sd02_column]],                  # 年龄 / Age
    Handedness = data[[sd20_column]],           # 利手 / Handedness
    Field_of_Study = data[[sd21_column]],       # 专业领域 / Field of study
    Material = data[[lg04_columns[i]]],         # 当前循环中的材料标识符 / Material identifier for the current iteration
    Realism_Score = data[[ci_columns[i]]],      # 当前材料的现实感评分 / Realism score for the current material
    Categorizing_Expressions_Score = if (i <= length(dc_columns)) data[[dc_columns[i]]] else NA,  # 表情分类评分（若存在） / Categorizing expressions score (if exists)
    Arousal_Score = if (i <= length(ra_columns)) data[[ra_columns[i]]] else NA   # 唤起评分（若存在） / Arousal score (if exists)
  )
  
  # 将临时数据框追加到对齐后的数据框中 / Append the temporary data frame to the aligned data frame
  aligned_data <- bind_rows(aligned_data, temp_df)
}

# 将整理后的数据保存为Excel文件 / Save the cleaned data as an Excel file
write.xlsx(aligned_data, "aligned_data.xlsx")

# 查看整理后的数据结构（可选） / View the structure of the cleaned data (optional)
head(aligned_data)

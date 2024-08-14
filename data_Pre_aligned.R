# 载入必要的库
install.packages("openxlsx")
library(dplyr)
library(openxlsx)  # 用于保存Excel文件
library(readxl)    # 用于读取Excel文件

# 读取数据
data <- read_excel("N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/data_JGFacialExpressionsRating_2024-08-14_13-54.xlsx")

# 提取相关列的名称
case_column <- "CASE"
lg02_column <- "LG02_01"  # 实验组别列
sd01_column <- "SD01"      # 性别列
sd02_column <- "SD02_01"   # 年龄列
sd20_column <- "SD20"      # 利手列
sd21_column <- "SD21_01"   # 专业领域列

lg04_columns <- grep("^LG04", names(data), value = TRUE)
ci_columns <- grep("^CI", names(data), value = TRUE)
dc_columns <- grep("^DC", names(data), value = TRUE)
ra_columns <- grep("^RA", names(data), value = TRUE)

# 初始化一个空的数据框来存储对齐后的数据
aligned_data <- data.frame()

# 对每一组 LG04 对应的 CI, DC, RA 列进行处理
for (i in seq_along(lg04_columns)) {
  temp_df <- data.frame(
    CASE = data[[case_column]],                 # 被试的编号
    Group = data[[lg02_column]],                # 实验组别
    Gender = data[[sd01_column]],               # 性别
    Age = data[[sd02_column]],                  # 年龄
    Handedness = data[[sd20_column]],           # 利手
    Field_of_Study = data[[sd21_column]],       # 专业领域
    Material = data[[lg04_columns[i]]],         # 当前循环中的材料标识符
    Realism_Score = data[[ci_columns[i]]],      # 当前材料的现实感评分
    Categorizing_Expressions_Score = if (i <= length(dc_columns)) data[[dc_columns[i]]] else NA,  # 表情分类评分（若存在）
    Arousal_Score = if (i <= length(ra_columns)) data[[ra_columns[i]]] else NA   # 唤起评分（若存在）
  )
  
  aligned_data <- bind_rows(aligned_data, temp_df)
}

# 将整理后的数据保存为Excel文件
write.xlsx(aligned_data, "aligned_data.xlsx")

# 查看整理后的数据结构（可选）
head(aligned_data)

# ---------------------------
# 完整脚本：去除背景灰度部分
# ---------------------------

# Step 1: 加载必要的 R 包
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(stringr)
library(patchwork)
library(cowplot)


# Step 2: 读取数据
data <- read_excel("N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx")

# Step 3: 数据预处理
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
    Model = str_extract(Material, "Fema\\d+|Male\\d+"),
    Model_Gender = case_when(
      str_detect(Material, "Fema") ~ "Female",
      str_detect(Material, "Male") ~ "Male"
    ),
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

# Step 4: 初始化 UHR 结果
uhr_results_all <- data.frame(
  Model = character(),
  Expression_Type = character(),
  UHR = numeric(),
  Chance_UHR = numeric(),
  Model_Gender = character(),
  stringsAsFactors = FALSE
)

# Step 5: 计算 UHR 和机会水平
calculate_uhr_chance_uhr <- function(conf_matrix) {
  uhr_results <- numeric(length(rownames(conf_matrix)))
  chance_uhr_results <- numeric(length(rownames(conf_matrix)))
  N <- sum(conf_matrix)
  for (i in seq_along(rownames(conf_matrix))) {
    emotion <- rownames(conf_matrix)[i]
    a <- conf_matrix[emotion, emotion]
    b <- sum(conf_matrix[emotion, ]) - a
    d <- sum(conf_matrix[, emotion]) - a
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

# Step 6: 计算每个模型的 UHR
models <- unique(data$Model)
expected_emotions <- c("Affiliation", "Disgust", "Dominance", "Enjoyment", "Neutral")
for (mod in models) {
  mod_data <- data %>% filter(Model == mod)
  conf_matrix <- mod_data %>%
    filter(Expression_Type != "Other", Chosen_Expression != "Other") %>%
    count(Expression_Type, Chosen_Expression) %>%
    tidyr::spread(key = Chosen_Expression, value = n, fill = 0) %>%
    tidyr::complete(Expression_Type = expected_emotions, fill = list(n = 0))
  conf_matrix_matrix <- as.matrix(conf_matrix[, -1])
  rownames(conf_matrix_matrix) <- conf_matrix$Expression_Type
  uhr_chance_results <- calculate_uhr_chance_uhr(conf_matrix_matrix)
  for (i in seq_along(expected_emotions)) {
    uhr_results_all <- rbind(uhr_results_all, data.frame(
      Model = mod,
      Expression_Type = expected_emotions[i],
      UHR = uhr_chance_results$UHR[i],
      Chance_UHR = uhr_chance_results$Chance_UHR[i],
      Model_Gender = first(mod_data$Model_Gender)
    ))
  }
}

# Step 7: 提取合理性得分并计算综合得分
uhr_results_all <- uhr_results_all %>% mutate(Performance_Above_Chance = UHR - Chance_UHR)
plausibility_col <- ifelse("Realism_Score" %in% colnames(data), "Realism_Score", "Plausibility_Score")
plausibility_scores <- data %>%
  group_by(Model) %>%
  summarise(Plausibility = mean(.data[[plausibility_col]], na.rm = TRUE), .groups = 'drop')
average_uhr_per_model <- uhr_results_all %>%
  group_by(Model, Model_Gender) %>%
  summarise(Average_UHR = mean(UHR, na.rm = TRUE), .groups = 'drop') %>%
  mutate(Arcsine_UHR = asin(sqrt(Average_UHR)))
combined_data <- average_uhr_per_model %>%
  left_join(plausibility_scores, by = "Model") %>%
  mutate(
    Z_Arcsine_UHR = scale(Arcsine_UHR),
    Z_Plausibility = scale(Plausibility),
    Combined_Score = (Z_Arcsine_UHR + Z_Plausibility) / 2
  ) %>%
  arrange(desc(Combined_Score)) %>%
  mutate(Index = row_number(), Group = ceiling(Index / 9))

# Step 8: 创建条形图
bar_plot <- ggplot(combined_data, aes(x = Combined_Score, y = reorder(Model, Combined_Score), fill = Model_Gender)) +
  geom_col(width = 0.7) +
  scale_fill_manual(values = c("Female" = "pink", "Male" = "steelblue")) +
  labs(x = "Combined Score", y = "Model", title = "Model Ranking by Combined Score", fill = "Gender") +
  theme_minimal() +
  theme(
    axis.text.y = element_text(size = 5),
    plot.title = element_text(size = 18, face = "bold")
  )

# Step 9: 创建分组饼图
pie_charts <- combined_data %>%
  group_split(Group) %>%
  lapply(function(group_data) {
    gender_counts <- group_data %>%
      group_by(Model_Gender) %>%
      summarise(n = n(), .groups = 'drop')
    ggplot(gender_counts, aes(x = "", y = n, fill = Model_Gender)) +
      geom_bar(stat = "identity", width = 1) +
      coord_polar(theta = "y") +
      scale_fill_manual(values = c("Female" = "pink", "Male" = "steelblue")) +
      labs(title = paste0("Group ", unique(group_data$Group), ": Ranks ",
                          min(group_data$Index), "-", max(group_data$Index))) +
      theme_void() +
      theme(
        plot.title = element_text(size = 10, face = "bold", hjust = 0.5),
        legend.position = "none"
      )
  })
pie_panel <- wrap_plots(pie_charts, ncol = 1)

# Step 10: 合并条形图和饼图
final_plot <- plot_grid(
  pie_panel, bar_plot,
  ncol = 2, rel_widths = c(1, 3)
) +
  plot_annotation(
    title = "Model Ranking by Combined Score with Group-wise Gender Distribution",
    subtitle = "Left: Pie charts for gender distribution per group; Right: Horizontal bar chart with ranking and group divisions",
    theme = theme(
      plot.title = element_text(size = 16, face = "bold"),
      plot.subtitle = element_text(size = 12)
    )
  )

# 保存图形
print(final_plot)
ggsave("Updated_Model_Ranking_Visualization.png", final_plot, width = 14, height = 10)

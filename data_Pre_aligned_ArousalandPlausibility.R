# ---------------------------------------------------------------------------
# Step 0: Install and load required packages (uncomment install commands if needed)
if(!require("readxl")) install.packages("readxl")
if(!require("dplyr")) install.packages("dplyr")
if(!require("lmerTest")) install.packages("lmerTest")
if(!require("emmeans")) install.packages("emmeans")
if(!require("ggplot2")) install.packages("ggplot2")
if(!require("car")) install.packages("car")

library(readxl)
library(dplyr)
library(lmerTest)
library(emmeans)
library(ggplot2)
library(car)

# ---------------------------------------------------------------------------
# Step 1: Read the data
file_path <- "N:/JinLab/Personal_JG_Lab/R_course/Facial Expressions Rating Task/aligned_data.xlsx"
data <- read_excel(file_path)

# ---------------------------------------------------------------------------
# Step 1.1: Convert rater gender (Rater_Gender) – assuming Gender column: 1 = Female, 2 = Male
data <- data %>% 
  mutate(Rater_Gender = case_when(
    Gender == 1 ~ "Female",
    Gender == 2 ~ "Male",
    TRUE ~ NA_character_
  ))

# ---------------------------------------------------------------------------
# Step 2: Extract expression type from the Material column
data <- data %>% 
  mutate(Expression_Type = case_when(
    grepl("dis", Material, ignore.case = TRUE) ~ "Disgust",
    grepl("enj", Material, ignore.case = TRUE) ~ "Enjoyment",
    grepl("aff", Material, ignore.case = TRUE) ~ "Affiliation",
    grepl("dom", Material, ignore.case = TRUE) ~ "Dominance",
    grepl("neu", Material, ignore.case = TRUE) ~ "Neutral",
    TRUE ~ "Other"
  ))

# ---------------------------------------------------------------------------
# Step 3: Map Categorizing_Expressions_Score to Chosen_Expression
data <- data %>% 
  mutate(Chosen_Expression = case_when(
    Categorizing_Expressions_Score == 1 ~ "Enjoyment",
    Categorizing_Expressions_Score == 2 ~ "Affiliation",
    Categorizing_Expressions_Score == 3 ~ "Dominance",
    Categorizing_Expressions_Score == 4 ~ "Disgust",
    Categorizing_Expressions_Score == 5 ~ "Neutral",
    Categorizing_Expressions_Score == 6 ~ "Other",
    TRUE ~ NA_character_
  ))

# ---------------------------------------------------------------------------
# Additional step: Create a Correct indicator (for potential future analyses)
# Mapping: Enjoyment = 1, Affiliation = 2, Dominance = 3, Disgust = 4, Neutral = 5
mapping <- c("Enjoyment" = 1, "Affiliation" = 2, "Dominance" = 3, "Disgust" = 4, "Neutral" = 5)
data <- data %>% 
  mutate(Correct = ifelse(mapping[Expression_Type] == Categorizing_Expressions_Score, 1, 0))

# ---------------------------------------------------------------------------
# Step 3.1: Extract other variables from Material
data <- data %>% 
  mutate(
    # Face_Gender: "fema" indicates Female, "male" indicates Male
    Face_Gender = case_when(
      grepl("fema", Material, ignore.case = TRUE) ~ "Female",
      grepl("male", Material, ignore.case = TRUE) ~ "Male",
      TRUE ~ NA_character_
    ),
    # Version: based on whether Material contains "L" or "R"
    Version = case_when(
      grepl("L", Material, ignore.case = TRUE) ~ "L",
      grepl("R", Material, ignore.case = TRUE) ~ "R",
      TRUE ~ NA_character_
    )
    # Group: directly use the Group variable from the original data
  )

# ---------------------------------------------------------------------------
# Step 4: Filter data – keep records where Expression_Type is not "Other"
data_filtered <- data %>% filter(Expression_Type != "Other")

# ---------------------------------------------------------------------------
# Step 5: Set Expression_Type as a factor with the specified order (using "Neutral" as the reference level)
expected_emotions <- c("Neutral", "Affiliation", "Disgust", "Dominance", "Enjoyment")
data_filtered <- data_filtered %>% 
  mutate(Expression_Type = factor(Expression_Type, levels = expected_emotions))

# ---------------------------------------------------------------------------
# Step 5.1: Calculate descriptive statistics by expression type for Arousal_Score and Realism_Score
desc_stats <- data_filtered %>%
  group_by(Expression_Type) %>%
  summarise(
    Mean_Arousal = mean(Arousal_Score, na.rm = TRUE),
    SD_Arousal   = sd(Arousal_Score, na.rm = TRUE),
    Mean_Realism = mean(Realism_Score, na.rm = TRUE),
    SD_Realism   = sd(Realism_Score, na.rm = TRUE),
    n = n()
  )
print(desc_stats)

# ---------------------------------------------------------------------------
# Step 6: LMM Analysis for Arousal_Score
lmm_arousal <- lmer(Arousal_Score ~ Expression_Type + Group + Rater_Gender + Face_Gender + Version + (1 | CASE),
                    data = data_filtered)
cat("===== LMM Arousal Model =====\n")
print(summary(lmm_arousal))
emms_arousal <- emmeans(lmm_arousal, ~ Expression_Type)
pairwise_arousal <- pairs(emms_arousal)
cat("\nArousal Model Pairwise Comparisons (Expression_Type):\n")
print(pairwise_arousal)

# ---------------------------------------------------------------------------
# Step 7: LMM Analysis for Realism_Score
lmm_realism <- lmer(Realism_Score ~ Expression_Type + Group + Rater_Gender + Face_Gender + Version + (1 | CASE),
                    data = data_filtered)
cat("\n===== LMM Realism Model =====\n")
print(summary(lmm_realism))
emms_realism <- emmeans(lmm_realism, ~ Expression_Type)
pairwise_realism <- pairs(emms_realism)
cat("\nRealism Model Pairwise Comparisons (Expression_Type):\n")
print(pairwise_realism)

# ---------------------------------------------------------------------------
# Step 8: Model Diagnostics – Check assumptions for the Arousal and Realism models

# Arousal model diagnostics
resid_arousal <- residuals(lmm_arousal)
fitted_arousal <- fitted(lmm_arousal)
standard_resid_arousal <- resid_arousal / sigma(lmm_arousal)

par(mfrow = c(2,2))
qqnorm(resid_arousal, main = "Arousal Model Q-Q Plot")
qqline(resid_arousal, col = "red")
hist(resid_arousal, breaks = 30, main = "Arousal Model Residual Histogram", xlab = "Residuals")
plot(fitted_arousal, resid_arousal, main = "Arousal: Fitted vs Residuals", xlab = "Fitted Values", ylab = "Residuals")
abline(h = 0, col = "red")
plot(fitted_arousal, standard_resid_arousal, main = "Arousal: Fitted vs Standardized Residuals", xlab = "Fitted Values", ylab = "Standardized Residuals")
abline(h = 0, col = "red")
par(mfrow = c(1,1))

# Realism model diagnostics
resid_realism <- residuals(lmm_realism)
fitted_realism <- fitted(lmm_realism)
standard_resid_realism <- resid_realism / sigma(lmm_realism)

par(mfrow = c(2,2))
qqnorm(resid_realism, main = "Realism Model Q-Q Plot")
qqline(resid_realism, col = "red")
hist(resid_realism, breaks = 30, main = "Realism Model Residual Histogram", xlab = "Residuals")
plot(fitted_realism, resid_realism, main = "Realism: Fitted vs Residuals", xlab = "Fitted Values", ylab = "Residuals")
abline(h = 0, col = "red")
plot(fitted_realism, standard_resid_realism, main = "Realism: Fitted vs Standardized Residuals", xlab = "Fitted Values", ylab = "Standardized Residuals")
abline(h = 0, col = "red")
par(mfrow = c(1,1))

# ---------------------------------------------------------------------------
# Step 9: Conduct Shapiro-Wilk and Levene tests on the residuals
if(length(resid_arousal) > 5000) {
  set.seed(123)
  shapiro_result_arousal <- shapiro.test(sample(resid_arousal, 5000))
} else {
  shapiro_result_arousal <- shapiro.test(resid_arousal)
}
cat("\nShapiro-Wilk Test (Arousal Model):\n")
print(shapiro_result_arousal)

levene_result_arousal <- leveneTest(resid_arousal ~ as.factor(fitted_arousal > median(fitted_arousal)))
cat("\nLevene Test for Homogeneity of Variance (Arousal Model):\n")
print(levene_result_arousal)

if(length(resid_realism) > 5000) {
  set.seed(123)
  shapiro_result_realism <- shapiro.test(sample(resid_realism, 5000))
} else {
  shapiro_result_realism <- shapiro.test(resid_realism)
}
cat("\nShapiro-Wilk Test (Realism Model):\n")
print(shapiro_result_realism)

levene_result_realism <- leveneTest(resid_realism ~ as.factor(fitted_realism > median(fitted_realism)))
cat("\nLevene Test for Homogeneity of Variance (Realism Model):\n")
print(levene_result_realism)

# ---------------------------------------------------------------------------
# Step 10: Save all results to a text file
sink("Arousal_Realism_LMM_Results.txt")
cat("===== Descriptive Statistics =====\n")
print(desc_stats)
cat("\n===== LMM Arousal Model =====\n")
print(summary(lmm_arousal))
cat("\nArousal Model Pairwise Comparisons (Expression_Type):\n")
print(pairwise_arousal)
cat("\n\n===== LMM Realism Model =====\n")
print(summary(lmm_realism))
cat("\nRealism Model Pairwise Comparisons (Expression_Type):\n")
print(pairwise_realism)
cat("\n\n===== Arousal Model Residual Tests =====\n")
cat("\nShapiro-Wilk Test (Arousal Model):\n")
print(shapiro_result_arousal)
cat("\nLevene Test (Arousal Model):\n")
print(levene_result_arousal)
cat("\n\n===== Realism Model Residual Tests =====\n")
cat("\nShapiro-Wilk Test (Realism Model):\n")
print(shapiro_result_realism)
cat("\nLevene Test (Realism Model):\n")
print(levene_result_realism)
sink()

cat("Analysis complete. Results saved to 'Arousal_Realism_LMM_Results.txt'\n")

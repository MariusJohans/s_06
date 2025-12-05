
##=============================================================================
##
## This code was written in  RStudio 2022.07.1+554 "Spotted Wakerobin"
## release. Functions and package compatibility may be subject to
## change.
##
##=============================================================================

## Libraries
pkgs <- c(
  "foreign","readr","lavaan","semPlot","psych","apaTables","memisc",
  "car","readxl","knitr","e1071","effsize","ltm","mice",
  "naniar","finalfit","writexl","semTools","MASS","simsem","simr",
  "ordinal","dplyr","tidyr","Hmisc","weights","ggplot2", "ggpubr"
)
for (p in pkgs) {
  if (!require(p, character.only = TRUE)) {
    install.packages(p)
    library(p, character.only = TRUE)
  }
}

Data <- read_excel("Data.xlsx", col_names = TRUE)

##=============================================================================
##
## Descriptive statistics
##
##=============================================================================

describe(Data$GRADE_NUM[Data$INT == "Y"])  
describe(Data$GRADE_NUM[Data$INT == "N"])   
shapiro.test(Data$GRADE_NUM[Data$INT == "Y"])
shapiro.test(Data$GRADE_NUM[Data$INT == "N"])

Data$GRADE <- factor(Data$GRADE, levels = c("F", "E", "D", "C", "B", "A"), 
                     ordered = TRUE)
Data$GRADE <- factor(Data$GRADE)
Data$INT   <- as.factor(Data$INT)
Data$YEAR  <- as.factor(Data$YEAR)

## Mann-Whitney U Test (difference before vs after)
Data$GRADE_NUM <- as.numeric(Data$GRADE)
wilcox_test <- wilcox.test(GRADE_NUM ~ INT, data = Data)
wilcox_test
## p = 0.003 => Significant effect of the intervention

##=============================================================================
##
## Plots
##
##=============================================================================

## The Freeman Plot
mean_data <- data.frame(
  INT = c("N", "Y"),
  Mean = tapply(Data$GRADE_NUM, Data$INT, mean)
  )
ggplot(Data, aes(x = GRADE_NUM, fill = INT, color = INT)) +
  geom_density(alpha = 0.3, bw = 0.65) +
  scale_x_continuous(
    breaks = 1:6,
    labels = c("F", "E", "D", "C", "B", "A"),
    name = "Grade") +
  scale_y_continuous(name = "Density") +
  labs(title = "Grade distributions before and after intervention") +
  geom_vline(
    data = mean_data,  
    aes(xintercept = Mean, color = INT),
    linetype = "dashed",
    size = 1) +
  scale_fill_manual(values = c("N" = "darkorange", "Y" = "cyan4"), 
                    labels = c("N" = "Before intervention", 
                               "Y" = "After intervention"),
                    name = NULL) +
  scale_color_manual(values = c("N" = "darkorange", "Y" = "cyan4"), 
                     labels = c("N" = "Before intervention", 
                                "Y" = "After intervention"),
                     name = NULL) +
  theme_minimal() +
  theme_bw() + theme(panel.border = element_blank(), 
                     panel.grid.major = element_blank(),
                       panel.grid.minor = element_blank(), 
                     axis.line = element_line(colour = "black"))

##=============================================================================
##
## Logistic regression model
##
##=============================================================================

model <- polr(Data$GRADE ~ Data$INT, data = Data, Hess = TRUE)
summary(model)
coefs <- coef(summary(model))
p_values <- pnorm(abs(coefs[, "t value"]), 
                  lower.tail = FALSE) * 2  ## Two-tailed test
coefs <- cbind(coefs, "p-value" = p_values)
print(coefs)

## Calculating CIs 
## Extract the coefficients and standard errors
coefficients <- summary(model)$coefficients[, "Value"]
std_errors   <- summary(model)$coefficients[, "Std. Error"]
## Calculate the 95% confidence intervals (Z = 1.96 for 95% CI)
z_value  <- 1.96
lower_CI <- coefficients - z_value * std_errors
upper_CI <- coefficients + z_value * std_errors
## Combine into a data frame
ci_table <-  data.frame(
  Estimate = coefficients,
  Lower_CI = lower_CI,
  Upper_CI = upper_CI
  )
print(ci_table)
## Exponentiate the CIs for odds ratios
lower_OR <- exp(lower_CI)
upper_OR <- exp(upper_CI)
## Combine into a data frame
or_ci_table <-  data.frame(
  OR_Estimate = exp(coefficients),  ## Odds ratios
  Lower_OR_CI = lower_OR,
  Upper_OR_CI = upper_OR
  )
print(or_ci_table)

## ANOVA
average_grades <- Data %>%
  group_by(YEAR) %>%
  summarise(avg_grade = mean(GRADE_NUM, na.rm = TRUE))
print(average_grades)
anova <- aov(GRADE_NUM ~ YEAR, data = Data)
summary(anova)
pairwise_res <- TukeyHSD(anova)
print(pairwise_res)
model <- lm(GRADE_NUM ~ YEAR, data = Data)
summary(model)

cliff.delta(GRADE ~ INT, data = Data)
## Kendall's Tau
kendall_result <- cor(Data$GRADE_NUM, Data$INT_num, method = "kendall")
print(kendall_result)
## Likelihood ratio test
model_baseline <- polr(GRADE ~ 1, data = Data, Hess = TRUE)
anova(model_baseline, model)

##=============================================================================
##
## Cumulative probability curves
##
##=============================================================================

grade_levels <- c("F", "E", "D", "C", "B", "A")
## Compute average predicted probabilities
## re-load model syntax above, polr looks outside the data frame too
prob_all  <- predict(model, type = "probs")
prob_post <- prob_all[Data$INT == "Y", ]
prob_pre  <- prob_all[Data$INT == "N", ]
avg_probs_pre  <- colMeans(prob_pre)
avg_probs_post <- colMeans(prob_post)
## Convert to cumulative probs
cum_avg_probs_pre  <- cumsum(avg_probs_pre)
cum_avg_probs_post <- cumsum(avg_probs_post)
## Put into a data frame
df_cum_probs <- data.frame(
  Grade = factor(grade_levels, levels = grade_levels),
  Pre = cum_avg_probs_pre,
  Post = cum_avg_probs_post
  ) %>%
  pivot_longer(cols = c("Pre", "Post"), names_to = "Group", 
               values_to = "CumulativeProbability")
ggplot(df_cum_probs, aes(x = Grade, y = CumulativeProbability, color = 
                           Group, group = Group)) +
  geom_point(size = 3) +
  #geom_line(linewidth = 1.2) +
  #geom_smooth(method = "loess", se = FALSE, linetype = 
  #"dashed", linewidth = 1, alpha = 0.5) +
  geom_smooth(method = "loess", se = FALSE, linewidth = 1, alpha = 0.5) +
  labs(
    #title = "Cumulative probability curve by intervention group",
    x = "Grade category",
    y = "Cumulative probability"
    ) +
  scale_color_manual(values = c("Pre" = "darkorange", "Post" = "cyan4")) +
  theme_minimal(base_size = 14)

##=============================================================================
##
## Power analysis using Monte Carlo
##
##=============================================================================

simulate_power <- function(n_sim = 1000, Data) {
  p_values <- numeric(n_sim)  ## Vector to store p-values
  ## Loop simulations
  for (i in 1:n_sim) {
    ## Resample the data (random sampling with replacement)
    sim_data <- Data[sample(1:nrow(Data), replace = TRUE), ]
    ## Fit the OLR model
    model_sim <- try(polr(GRADE ~ INT, data = sim_data, Hess = TRUE), 
                     silent = TRUE)
    ## Check if model fitting was successful
    if (inherits(model_sim, "try-error")) {
      next  ## Skip to next simulation if model fitting fails
      }
    ## Extract summary from the model
    summary_model <- summary(model_sim)
    ## Check the structure of summary_model$coefficients 
    print(summary_model$coefficients)
    ## Extract p-value for the INT variable
    try({
      p_value <- summary_model$coefficients["INT", "Pr(>|z|)"]
      }, silent = TRUE)
        if (exists("p_value") && !is.null(p_value)) {
            p_values[i] <- p_value
            }
      }
  ## Calculate power - proportion of simulations where p-value < 0.05
  power <- mean(p_values < 0.05, na.rm = TRUE)
  return(power)
}
## Run power simulation
power_estimate <- simulate_power(n_sim = 1000, Data = Data)
print(paste("Estimated Power:", power_estimate * 100, "%"))

##=============================================================================
##
## END
##

##=============================================================================

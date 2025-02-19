  # Set working directory

# Load required libraries
library(openxlsx)  # For reading Excel files
library(dplyr)     # For data manipulation
library(ggplot2)   # For visualization

# Read data from Excel sheets
data <- read.xlsx("../data/data_Bark.xlsx", 1)  # First sheet containing main data
data.3 <- read.xlsx("../data/data_Bark.xlsx", 2)  # Second sheet containing height data

#######################
########Method-1#######
#######################

# Convert Diameter from cm to meters
data$dbh_m <- data$Diameter.in.cm / 100

# Convert Double Bark Thickness from mm to meters
data$bt_m <- data$Double.Bark.thickness.in.mm / 1000

# Calculate bark volume using a cylindrical approximation formula
data$barkVol <- pi * data$Height.in.m * (
  (data$dbh_m / 2 + data$bt_m / 2)^2 - (data$dbh_m / 2)^2
)

# Merge with height data using Tree ID
data <- merge(data.3, data, by = 'Tree.ID')

# Get last observation for each Tree ID (highest recorded height)
last_obs <- data %>%
  group_by(Tree.ID) %>%
  filter(Height.in.m == max(Height.in.m)) %>%
  ungroup()

# Merge last observation data with original dataset
data <- merge(data, last_obs, by = "Tree.ID", suffixes = c("", "_last"))

# Calculate bark volume from the last height observation
data$last.vol <- pi * (data$Tree.height.in.m - data$Height.in.m_last) * (
  (data$dbh_m_last / 2 + data$bt_m_last / 2)^2 - (data$dbh_m_last / 2)^2
)

# Calculate Relative Height (RH%) as a percentage
data$RH <- (data$Height.in.m / data$Tree.height.in.m) * 100

# Compute Carbon Content (C%) using an empirical equation
data$C_percent <- 49.564 + 5.306^(-4.659 * data$RH)

# Compute Bark Fissure Index (BFI) using a regression equation
beta_0 <- 1.070
beta_1 <- -0.076
beta_2 <- 11.359
data$BFI <- beta_0 + beta_1 * (1 - exp(-beta_2 * data$RH))

# Compute Total Bark Volume (sum of last volume and initial volume)
data$total_bark_volume <- data$last.vol + data$barkVol

# Calculate Bark Carbon using total volume, carbon content, and BFI
data$bark_carbon <- data$total_bark_volume * (data$C_percent / 100) * data$BFI * 270

# Summarize results for each Tree ID
data.method1 <- data %>%
  group_by(Tree.ID) %>%
  summarise(
    dbh = dbh_m[Height.in.m == 1.3],  # Select DBH at height 1.3m
    h = first(Tree.height.in.m),  # Take the tree height
    bark.carbon = sum(bark_carbon, na.rm = TRUE),  # Total bark carbon
    bark.volume = sum(total_bark_volume, na.rm = TRUE),  # Total bark volume
    region = first(region)  # Assign region for each Tree ID
  )

###################################################################################################################
#######################
########Method-2#######
#######################

# Read data again for second method
data.2 <- read.xlsx("../data/data_Bark.xlsx", 1)
data.3 <- read.xlsx("../data/data_Bark.xlsx", 2)

# Merge datasets using Tree ID
data.new <- merge(data.2, data.3, by = 'Tree.ID')

# Assign values for Sinan et al. method
data.new$dbh <- ifelse(data.new$Height.in.m == 1.3, data.new$Diameter.in.cm, 0)
data.new$hst <- 0.3  # Stump height
data.new$dst <- ifelse(data.new$Height.in.m == 0.3, data.new$Diameter.in.cm, 0)  # Diameter at stump height
data.new$hlc <- data.new$crone.height  # Crown base height
data.new$h <- data.new$Tree.height.in.m  # Total tree height

### Extract last DBH recorded for each tree
last_dbh <- data.new %>%
  group_by(Tree.ID) %>%
  slice_max(order_by = Height.in.m, n = 1) %>%
  select(Tree.ID, Diameter.in.cm) %>%
  rename(dhk = Diameter.in.cm)  # Rename for clarity

# Merge DHK into the main dataset
data.new <- left_join(data.new, last_dbh, by = "Tree.ID")

### Adjustments for bark thickness calculations
data.new$bt.under <- ifelse(data.new$Height.in.m == 1.3, data.new$Double.Bark.thickness.in.mm, 0)

# Extract last bark thickness for each tree
last_bark <- data.new %>%
  group_by(Tree.ID) %>%
  slice_max(order_by = Height.in.m, n = 1) %>%
  select(Tree.ID, Double.Bark.thickness.in.mm) %>%
  mutate(bt.over = Double.Bark.thickness.in.mm) %>%
  select(Tree.ID, bt.over)

# Merge last bark thickness into dataset
data.new <- left_join(data.new, last_bark, by = "Tree.ID")

########### Summarize Key Variables
data.method2 <- data.new %>%
  group_by(Tree.ID) %>%
  summarise(
    dbh = max(dbh / 100),  
    hst = max(hst),  
    dst = max(dst / 100),  
    hlc = max(hlc),  
    h = max(h, na.rm = TRUE),      
    dhk = max(dhk / 100, na.rm = TRUE),  
    bt_under = max(bt.under / 1000, na.rm = TRUE),  
    bt.over = max(bt.over / 1000, na.rm = TRUE),  
    region = first(region)
  )

### Volume Calculations Based on Different Geometries

# Compute volume for different trunk sections
data.method2$vol.cylinder <- pi * (data.method2$dst)^2 / 4 * data.method2$hst
data.method2$ao <- pi * (data.method2$dbh)^2 / 4  
data.method2$au <- pi * (data.method2$dst)^2 / 4  

# Truncated Neiloid Volume
data.method2$vol.neloid <- (1.3 - data.method2$hst) / 4 * (
  data.method2$ao + data.method2$au + 
    (data.method2$ao^2 * data.method2$au)^0.333 + (data.method2$ao * data.method2$au^2)^0.333
)

# Truncated Cone Volume
data.method2$vol.cone <- pi * (data.method2$hlc - 1.3) / 3 * (
  data.method2$dhk^2 / 4 + data.method2$dbh^2 / 4 + (data.method2$dhk / 2) * (data.method2$dbh / 2)
)

# Paraboloid Volume
data.method2$vol.parab <- data.method2$dhk^2 / 4 * pi * (data.method2$h - data.method2$hlc) / 2.5

# Compute total bark volume
data.method2$bottom.vol <- data.method2$vol.cylinder + data.method2$vol.neloid + data.method2$vol.cone
data.method2$top.vol <- data.method2$vol.parab

data.method2$barkvol <- data.method2$bottom.vol * data.method2$bt_under + data.method2$top.vol * data.method2$bt.over
data.method2$bark.C <- data.method2$barkvol * 0.9 * 395

# Remove unwanted coloumns
data.method2<-data.method2[-c(11:18)]


# Combine datasets for comparison
data.compare <- bind_rows(
  data.method1 %>% select(Tree.ID, region, bark.carbon) %>% mutate(Method = "Hollow cylinder model"),
  data.method2 %>% select(Tree.ID, region, bark.C) %>% rename(bark.carbon = bark.C) %>% mutate(Method = "3D Geometrical model ")
)
##Boxplot
boxplot(bark.carbon ~ Method, 
        data = data.compare,
        border = "black",  # Set border color
        ylab = "Bark Carbon (kg)", 
        xlab = "", 
        main = "",
        ylim = c(0, 200))  # Set y-axis limits

# Add mean points
means <- tapply(data.compare$bark.carbon, data.compare$Method, mean, na.rm = TRUE)
points(1:length(means), means, col = "red", pch = 16, cex = 1.2)



t_test_result <- t.test(data.method1$bark.carbon, data.method2$bark.C, paired = TRUE)
print(t_test_result)



# Combine datasets for regional comparison
data.compare <- bind_rows(
  data.method1 %>% select(Tree.ID, region, bark.carbon) %>% mutate(Method = "Hollow Cylinder"),
  data.method2 %>% select(Tree.ID, region, bark.C) %>% rename(bark.carbon = bark.C) %>% mutate(Method = "3D Geometrical Model")
)

# Boxplot
ggplot(data.compare, aes(x = region, y = bark.carbon, fill = Method)) +
  geom_boxplot(alpha = 0.7, outlier.shape = NA, color = "black") +  # Boxplot with black outline
  stat_summary(fun = mean, geom = "point", shape = 16, size = 2, color = "red", position = position_dodge(width = 0.75)) + # Mean points
  scale_fill_manual(values = c("Hollow Cylinder" = "gray", "3D Geometrical Model" = "black")) +  # Custom colors
  labs(title = "",
       x = "",
       y = "Bark Carbon (kg)",
       fill = "Methods") +
  ylim(0, 200) +
  theme_minimal() +
  theme(
    panel.grid.major = element_blank(),  # Remove major grid
    panel.grid.minor = element_blank(),  # Remove minor grid
    panel.border = element_rect(color = "black", fill = NA, size = 0.7),  # Add outer border
    legend.background = element_rect(color = "black", fill = "white", size = 0.5),  # Add border around legend
    legend.position = c(0.24, 1),  # Move legend inside plot (adjust coordinates as needed)
    legend.justification = c(1, 1),  # Anchor legend to top-right
    axis.text.x = element_text(angle = 45, hjust = 1)  # Rotate x-axis labels
  )


# Run paired t-tests for each region
t_test_results <- data.compare %>%
  group_by(region) %>%
  summarise(
    t_test = list(t.test(
      bark.carbon[Method == "Hollow Cylinder"], 
      bark.carbon[Method == "3D Geometrical Model"], 
      paired = TRUE
    ))
  )

# Print results for each region
t_test_results$t_test

# Combine datasets for height-based analysis
data.compare.height <- bind_rows(
  data.method1 %>% select(Tree.ID, region, h, bark.carbon) %>% mutate(Method = "Hollow Cylinder"),
  data.method2 %>% select(Tree.ID, region, h, bark.C) %>% rename(bark.carbon = bark.C) %>% mutate(Method = "3D Geometrical Model")
)

# Scatter plot with trend line and legend inside the plot
ggplot(data.compare.height, aes(x = h, y = bark.carbon, color = Method)) +
  geom_point(alpha = 0.6, size = 2) +  # Scatter points
  geom_smooth(method = "loess", se = FALSE, linewidth = 1) +  # LOESS smooth curve
  labs(title = "",
       x = "Tree Height (m)",
       y = "Bark Carbon (kg)",
       color = "Methods")  +
  ylim(0, 200) +
  theme_minimal() +
  theme(
    legend.position = c(1, 1),  # Move legend inside plot (top-right)
    legend.justification = c(1, 1),  # Anchor legend to top-right corner
    legend.background = element_rect(fill = "white", color = "black", size = 0.5),  # Add border to legend
    panel.border = element_rect(color = "black", fill = NA, size = 0.7)  # Add border around plot
  )

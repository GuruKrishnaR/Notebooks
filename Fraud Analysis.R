install.packages("readxl")
install.packages("dplyr")
install.packages("ggplot2")
install.packages("tidyr")
install.packages("ggdark")
library(readxl)
library(dplyr)
library(ggplot2)
library(tidyr)
library(ggplot2)

excel_sheets('cw_r.xlsx')

#Import the tables
fraud_types_cohort <- read_excel('cw_r.xlsx', sheet = 'Table 3d')
regions_cohort <- read_excel('cw_r.xlsx', sheet = 'Table 5')
personal_characteristics_cohort <- read_excel('cw_r.xlsx', sheet = 'Table 7')
household_and_area_cohort <- read_excel('cw_r.xlsx', sheet = 'Table 8')

#Slicing
df_Banking_and_credit <- fraud_types_cohort %>%
  slice(10:14) %>%
  select(1:11) 

#Checking for the column names
print(colnames(df_Banking_and_credit))

#Slicing the desired row for column names
colnames_BnC <- slice(fraud_types_cohort,8)
colnames(df_Banking_and_credit) <- colnames_BnC

#Visualizing the dataframe in current form
print(df_Banking_and_credit)

#Head function restricted to 3 rows
head(df_Banking_and_credit, 3)

#Displaying the structure of the dataframe
str(df_Banking_and_credit)

#Finding Dispersion using columns
dispersion_BnC <- apply(df_Banking_and_credit[,-1],2,sd)
print(dispersion_BnC)

#Finding Central Tendency using columns
mean_BnC <- apply(df_Banking_and_credit[, -1], 2, mean)
print(mean_BnC)

warnings()

#Checking non-numeric characters in the dataframe
print(sapply(df_Banking_and_credit[, -1], class))

# Replace non-numeric characters and convert to numeric
df_Banking_and_credit[, -1] <- sapply(df_Banking_and_credit[, -1], function(x) as.numeric(gsub("[^0-9.]", "", x)))

mean_BnC <- apply(df_Banking_and_credit[, -1], 2, mean)
print(mean_BnC)

#Dispersion (Standard Deviation) & Mean (Central Tendency)
plot_BnC <- data.frame(Year = names(mean_BnC),Mean_Value = mean_BnC,Standard_Deviation = dispersion_BnC)
plot_BnC_Both <- tidyr::gather(plot_BnC, key = "Category", value = "Offences", -Year)
ggplot(plot_BnC_Both, aes(x = Year, y = Offences, fill = Category)) +
  geom_bar(stat = "identity", position = "stack", alpha = 0.7) +
  scale_y_continuous(labels = scales::comma) +
  labs(title = "Mean and Standard Deviation for Each Year - Banking & Credit",x = "Year",y = "Offences",fill = "Category") +
  theme_minimal()

#Slicing Consumer and retail frauds as df_Consumer_and_retail
df_Consumer_and_retail <- fraud_types_cohort %>%
  slice(40:46) %>%
  select(1:11)

#Checking for the column names
print(colnames(df_Consumer_and_retail))

#Checking for the column names
colnames_CnR <- slice(fraud_types_cohort,8)

#Assigning new column names inside the dataframe
colnames(df_Consumer_and_retail) <- colnames_CnR

#Visualizing the dataframe in current form
print(df_Consumer_and_retail)

#Head function restricted to 3 rows
head(df_Consumer_and_retail, 3)

#Displaying the structure of the dataframe
str(df_Consumer_and_retail)

#Finding Dispersion using columns
dispersion_CnR <- apply(df_Consumer_and_retail[,-1],2,sd)
print(dispersion_CnR)

#Finding Central Tendency using columns
mean_CnR <- apply(df_Consumer_and_retail[, -1], 2, mean)
print(mean_CnR)


warnings()

#Checking non-numeric characters in the dataframe
print(sapply(df_Consumer_and_retail[, -1], class))

# Replace non-numeric characters and convert to numeric
df_Consumer_and_retail[, -1] <- sapply(df_Consumer_and_retail[, -1], function(x) as.numeric(gsub("[^0-9.]", "", x)))

#Central Tendency using Mean function
mean_CnR <- apply(df_Consumer_and_retail[, -1], 2, mean)
print(mean_CnR)


#Dispersion (Standard Deviation) & Mean (Central Tendency)
plot_CnR <- data.frame(Year = names(mean_CnR),Mean_Value = mean_CnR,Standard_Deviation = dispersion_CnR)
plot_CnR_Both <- tidyr::gather(plot_CnR, key = "Category", value = "Offences", -Year)
ggplot(plot_CnR_Both, aes(x = Year, y = Offences, fill = Category)) +
  geom_bar(stat = "identity", position = "stack", alpha = 0.7) +
  scale_y_continuous(labels = scales::comma) +
  labs(title = "Mean and Standard Deviation for Each Year - Consumer & Retail",x = "Year",y = "Offences",fill = "Category") +
  theme_minimal()

#region_rows<-c(12,16,22,27,33,38,45,48,54,60)
regions_split<-regions_cohort %>%
  slice(c(12,16,22,27,33,38,45,48,54,60)) %>%
  select(2:4)

print(regions_split)

colnames(regions_split) <- c("Area_Name", "Number_of_Offences", "Rate_per_1000_Population")

print(colnames(regions_split))

print(regions_split)

regions_split$Number_of_Offences <- as.numeric(regions_split$Number_of_Offences)
regions_split$Rate_per_1000_Population <- as.numeric(gsub("[^0-9.]", "", regions_split$Rate_per_1000_Population))
regions_split$Area_Name <- factor(regions_split$Area_Name, levels = regions_split$Area_Name[order(regions_split$Number_of_Offences)])

print(regions_split)

install.packages("ggdark")
library(ggplot2)


# Create a ggplot with points and color-coded by Rate per 1,000 Population
p <- ggplot(regions_split, aes(x = Number_of_Offences, y = Area_Name, color = Rate_per_1000_Population)) +  geom_point() +  scale_color_gradient(low = "lightyellow", high = "red") +  labs(
    title = "Number of Offences by Area - Regions",    x = "Number of Offences",    y = "Area Name",    color = "Rate per 1,000 Population"  ) +  theme_dark()

# Display the plot
print(p)


#Slice the rows needed
county_rows<-c(13:15,17:21,23:26,28:32,34:37,39:44,46,47,49:53,55:59,61:64) 
county_split<-regions_cohort %>% slice(county_rows) %>% select(2:4)

print(colnames(county_split))

#Rename columns to remove spaces
colnames(county_split) <- c("Area_Name", "Number_of_Offences", "Rate_per_1000_Population")

print(county_split)

county_split$Number_of_Offences <- as.numeric(county_split$Number_of_Offences)
county_split$Rate_per_1000_Population <- as.numeric(gsub("[^0-9.]", "", county_split$Rate_per_1000_Population))
county_split$Area_Name <- factor(county_split$Area_Name, levels = county_split$Area_Name[order(county_split$Number_of_Offences)])

#county_split<-county_split %>%
 # arrange(`Number of offences`) %>%


install.packages("ggdark")
library(ggplot2)


# Create a ggplot with points and color-coded by Rate per 1,000 Population
p <- ggplot(county_split, aes(x = Number_of_Offences, y = Area_Name, color = Rate_per_1000_Population)) +
  geom_point() + 
  scale_color_gradient ( low = "lightyellow", high = "red" ) +
  labs(title = "Number of Offences by Area - Counties", x = "Number of Offences", y = "Area Name", color = "Rate per 1,000 Population") +
  theme_dark()

# Display the plot
print(p)


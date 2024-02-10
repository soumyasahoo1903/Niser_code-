import pandas as pd
import numpy as np
from scipy import stats

# Read the Excel file
data = pd.read_excel("F:/NISER internship/microbe and metbolite data analysis code/Host_metabolome/week_3_avg1.xlsx")

# Calculate and add the sample size column
data['Sample Size'] = data.iloc[:, 1:-1].count(axis=1)

# Eliminate rows with Sample Size less than 2
data = data[data['Sample Size'] >= 2]

# Calculate the mean abundance column
data['Average'] = data.mean(axis=1)

# Calculate the standard deviation column
data['Standard Deviation'] = data.iloc[:, 1:-1].std(axis=1)

# Set the desired confidence level
confidence_level = 0.95

# Calculate the confidence interval column
margin_of_error = data['Standard Deviation'] / np.sqrt(data['Sample Size'])
z_score = np.abs(stats.norm.ppf((1 - confidence_level) / 2))
data['Confidence Interval'] = margin_of_error * z_score * 2

# Calculate the sample variance column
data['CV'] = data.iloc[:, 1:].var(axis=1, ddof=1)

# Save the updated DataFrame to a new Excel file
data.to_excel("F:/NISER internship/microbe and metbolite data analysis code/Host_metabolome/week_3_avg1.xlsx", index=False)

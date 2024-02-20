import pandas as pd
from docxtpl import DocxTemplate
import sys

# Load the Excel sheet into a DataFrame
df = pd.read_excel('Flintridge_Responses.xlsx', sheet_name='Miles_&_Steps_Per_Day_Flintridg', header=0)

# print(df.head())  # Display the first few rows of the DataFrame
# print(df.columns.to_list())  # Display the column names

doc = DocxTemplate("template.docx")

# Command Line arguments for which template to use.
arguments = sys.argv

script_name = arguments[0]  # default, subsequent args are arguments[i]

# Then incapsulate in a loop to do this every time and create the first five users documents.

# Summary statistics. Calculate from the data frame elements

Num_Days = 30 #This should be calculated from the number of observations
Total_Miles = round(df.loc[1, 'Cum (Miles)'], 2)
Total_Steps = round(df.loc[1, 'Cum (Steps)'], 2)
Avg_Miles = round(Total_Miles/Num_Days, 2)
Avg_Steps = round(Total_Steps/Num_Days, 2)
Miles_12_Months = round(Total_Miles*12, 2)

best_miles = 0
best_miles_day = 0
best_steps = 0
best_steps_day = 0

# getting the different variables into a better format
for i in range(5, df.shape[1]):
    # if it is odd, it is the miles
    if (i % 2):
        if (df.iloc[1][i] >= best_miles):
            best_miles = round(df.iloc[1][i], 2)
            best_miles_day = df.columns[i]
    # if it is even, it is the steps
    else:
        if (df.iloc[1][i] >= best_steps):
            best_steps = round(df.iloc[1][i], 2)
            best_steps_day = df.columns[i-1]

# Convert the column header from date/time to MMM DDD
print(type(best_miles_day))
print(type(best_steps_day))
# Extract # print(month)
# print(day)

# Get all of the right variables in the context using loc.

context = {
    'First' : df.loc[1, 'First Name'],
    'Last' : df.loc[1, 'Last Name'],
    'School' : 'Flintridge Prepatory School',
    'Total_Miles' : Total_Miles,
    'Avg_Miles_Per_Day' : Avg_Miles,
    'Best_Miles' : best_miles,
    'Best_Steps' : best_steps,
    'Best_Miles_Day' : best_miles_day,
    'Total_Steps' : Total_Steps,
    'Avg_Steps_Per_Day' : Avg_Steps,
    'Best_Steps_Day' : best_steps_day,
    'Miles_12_Months' : Miles_12_Months
    }

# This is the save version of the document after the context
doc.render(context)
doc.save('reports/template_Rendered.docx')
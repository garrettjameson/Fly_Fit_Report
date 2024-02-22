import pandas as pd
from docxtpl import DocxTemplate, InlineImage
import sys
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

# Load the Excel sheet into a DataFrame
df = pd.read_excel('Flintridge_Responses.xlsx', sheet_name='Miles_&_Steps_Per_Day_Flintridg', header=0)

# print(df.head())  # Display the first few rows of the DataFrame
# print(df.columns.to_list())  # Display the column names

# doc = DocxTemplate("templates/template.docx")
doc = DocxTemplate("workingtemplate.docx")

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
Miles_12_Months = int(Total_Miles*12)

Best_Miles = 0
Best_Miles_day = 0
Best_Steps = 0
Best_Steps_day = 0

# dictionary to hold date and mileage
miles_data = {}

# getting the different variables into a better format
for i in range(5, df.shape[1]):
    # if it is odd, it is the miles
    if (i % 2):
        # extract the miles for a barchart!
        miles_data[df.columns[i]] = df.iloc[1][i]

        if (df.iloc[1][i] >= Best_Miles):
            Best_Miles = round(df.iloc[1][i], 2)
            Best_Miles_day = df.columns[i]
    # if it is even, it is the steps
    else:
        if (df.iloc[1][i] >= Best_Steps):
            Best_Steps = round(df.iloc[1][i], 2)
            Best_Steps_day = df.columns[i-1]

def barchartMaker(data):
    # make the bar chart
    # assumes data is in a dictionary
    dates = list(data.keys())
    miles = list(data.values())
    
    # create labels from Data for axis
    day_nums = [str(obs.day) for obs in dates]
    months = [obs.strftime('%B') for obs in dates]

    fig, ax1 = plt.subplots()
    # cap size
    fig.set_size_inches(4.9, 2)

    # Create the bar chart
    bars = ax1.bar(day_nums, miles)

    # title
    ax1.set_title('Your Daily Miles')

    # x-axis ticks adjusted
    ax1.set_xticks(day_nums[::2])
    ax1.xaxis.set_ticks_position('bottom')
    ax1.yaxis.set_major_locator(MultipleLocator(5))

    # add additional x axis with month
    ax2 = ax1.twiny()
    ax2.set_xlim(ax1.get_xlim())
    ax2.set_xticks([ax2.get_xticks()[0], ax2.get_xticks()[-1]])
    ax2.xaxis.set_ticklabels(list(dict.fromkeys(months)))
    ax2.xaxis.set_ticks_position('bottom')
    ax2.spines['bottom'].set_position(('outward', 20))

    # Add a horizontal line at y = 5
    plt.axhline(y=5, color='black', linestyle='--')

    # Set color based on the condition (less than 5)
    for bar, height in zip(bars, miles):
        if height < 5:
            bar.set_color('grey')  # Set color to grey for bars less than 5
        else:   
            bar.set_color('dodgerblue')  # Set color to green for bars greater than or equal to 5


    plt.tight_layout()  # Adjust layout to prevent clipping of labels
    # plt.show()
    plt.savefig('individual_plot.png')
    plt.close()

barchartMaker(miles_data)

def linechartMaker(data):
    dates = list(data.keys())
    miles = list(data.values())

    # create labels from Data for axis
    day_nums = [str(obs.day) for obs in dates]
    months = [obs.strftime('%B') for obs in dates]
    # build the plot
    plt.figure(figsize=(6.5, 1.78))
    plt.plot(day_nums, miles)
    

    # adjust the axes
    plt.ylim(min(miles), max(miles))
    plt.xlim(min(day_nums), plt.xlim()[1])
    plt.xticks(day_nums[::2])
    plt.title('Team Miles')

    plt.fill_between(day_nums, miles, color='lightblue', alpha = 0.7)
    plt.tight_layout()
    plt.savefig('team_plot.png')
    plt.close()
    # plt.show()

linechartMaker(miles_data)

# Create the inline images for docxtpl using the saved images
Individual_Chart_Image = InlineImage(doc, "individual_plot.png")
Team_Chart_Image = InlineImage(doc, "team_plot.png")

# Convert the column header from date/time to MMM DDD
# Extract the month and day from best_xx_day
def MonthDayExtractor(date):
    month = date.strftime('%B')
    day = date.day

    # Convert day to have "st", "nd", "rd", or "th" suffix
    if (day in (1, 21, 31)):
        suffix = "st"
    elif (day in (2, 22)):
        suffix = "nd"
    elif (day in (3, 23)):
        suffix = "rd"
    else:
        suffix = "th"
    
    formatted_date = f"{month} {day}{suffix}"

    return formatted_date

Best_Steps_Date = MonthDayExtractor(Best_Steps_day)
Best_Miles_Date = MonthDayExtractor(Best_Miles_day)

# Get all of the right variables in the context using loc.

context = {
    'First' : df.loc[1, 'First Name'],
    'Last' : df.loc[1, 'Last Name'],
    'School' : 'Flintridge Prepatory School',
    'Total_Miles' : Total_Miles,
    'Avg_Miles_Per_Day' : Avg_Miles,
    'Best_Miles' : Best_Miles,
    'Best_Steps' : Best_Steps,
    'Best_Miles_Day' : Best_Miles_Date,
    'Total_Steps' : Total_Steps,
    'Avg_Steps_Per_Day' : Avg_Steps,
    'Best_Steps_Day' : Best_Steps_Date,
    'Miles_12_Months' : Miles_12_Months,
    'Individual_Plot_Image' : Individual_Chart_Image,
    'Team_Plot_Image' : Team_Chart_Image
    }

# This is the save version of the document after the context
doc.render(context)
doc.save('reports/template_Rendered.docx')
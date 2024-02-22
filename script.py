import pandas as pd
from docxtpl import DocxTemplate, InlineImage
import sys
import os
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import datafetch
from datetime import datetime

# Command Line arguments for which template to use.
arguments = sys.argv

script_name = arguments[0]  # default, subsequent args are arguments[i]
if len(arguments)>2:
    schoolName = arguments[1]
    teamName = arguments[2]
else:
    print('Must use arguments with school and then team name. ex script.py "Flintridge Prepatory School" "Hot Steppas"')
    exit()

df = datafetch.get_challenge_report_data("060a4971-a44c-4de7-8c75-0f777adf1a1a")

# doc = DocxTemplate("templates/template.docx")
doc = DocxTemplate("workingtemplate.docx")

def barchartMaker(dailyData):
    # make the bar chart
    observations = [list(obs.values()) for obs in dailyData]
    dates = [value[0] for value in observations]
    miles = [value[1] for value in observations]
    
    
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

    # adjust ylim to be 10% bigger than max
    if miles:
        ax1.set_xlim(0, max(miles) * 1.1)

    # add additional x axis with month
    # ax2 = ax1.twiny()
    # ax2.set_xlim(ax1.get_xlim())
    # ax2.set_xticks([ax2.get_xticks()[0], ax2.get_xticks()[-1]])
    # ax2.xaxis.set_ticklabels(months)
    # ax2.xaxis.set_ticks_position('bottom')
    # ax2.spines['bottom'].set_position(('outward', 20))

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
    plt.savefig(f'individual_plot.png')
    plt.close()

def linechartMaker(data):
    # assumes a dictionary of data
    dates = list(set(data.keys()))
    miles = list(data.values())
    sorted_dates = sorted(dates)

    # Calculate cumulative sum of values
    cumulative_miles = []
    running_sum = 0
    for date in sorted_dates:
        running_sum += data[date]
        cumulative_miles.append(running_sum)

    # create labels from Data for axis
    day_nums = [str(obs.day) for obs in dates]
    months = [obs.strftime('%B') for obs in dates]
    # build the plot
    plt.figure(figsize=(6.5, 1.78))
    plt.plot(day_nums, cumulative_miles, marker='o')
    

    # adjust the axes
    if cumulative_miles:
        plt.ylim(0, max(cumulative_miles) * 1.1)
        plt.xlim(min(day_nums), plt.xlim()[1])
        plt.xticks(day_nums[::2])
    plt.title('Cumulative Team Miles')

    plt.fill_between(day_nums, cumulative_miles, color='lightblue', alpha = 0.7)
    plt.tight_layout()
    # plt.show()
    plt.savefig('team_plot.png')
    plt.close()


# Create the inline images for docxtpl using the saved image
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

df['endTimestampTime'] = datetime.strptime(df['endTimestamp'], '%Y-%m-%dT%H:%M:%S.%fZ')
df['startTimestampTime'] = datetime.strptime(df['startTimestamp'], '%Y-%m-%dT%H:%M:%S.%fZ')

Num_Days = df['endTimestampTime']-df['startTimestampTime'] #This should be calculated from the number of observations

# Convert date strings to datetime objects for each participant
for participant in df['participants']:
    for daily_data in participant['dailyData']:
        daily_data['date'] = datetime.strptime(daily_data['date'], '%Y-%m-%dT%H:%M:%S.%fZ')


# dictionary to hold date mileage pairings
teamMilesData = {}

# loop to go through each participant
for participant in df['participants']:
    # loop through each observation
    for obs in participant['dailyData']:
        # check to see if the date is already in the dictionary.
        if obs['date'] in teamMilesData:
        # if yes, add the mileage to the current value.
            teamMilesData[obs['date']] = teamMilesData[obs['date']] + obs['mileage']
        else:
            # if not, add it as a key and add the mileage as a value
            teamMilesData[obs['date']] = obs['mileage']

# Make the Team Line Chart
linechartMaker(teamMilesData)

# THIS IS WHERE THE MAIN LOOP STARTS!
for participant in df['participants']:
    # calculating my own statistics
    individualData = participant
    Best_Miles = 0
    Best_Steps = 0
    for obs in individualData['dailyData']:
        if obs['mileage'] >= Best_Miles:
            Best_Miles = obs['mileage']
            Best_Miles_Day = obs['date']
        if obs['steps'] >= Best_Steps:
            Best_Steps = obs['steps']
            Best_Steps_Day = obs['date']

    Best_Miles_Date = MonthDayExtractor(Best_Miles_Day)
    Best_Steps_Date = MonthDayExtractor(Best_Steps_Day)

    # call for individual barchart
    barchartMaker(individualData['dailyData'])
    Individual_Chart_Image = InlineImage(doc, "individual_plot.png")

    # Get all of the right variables in the context using loc.

    context = {
        # for the header
        'School' : schoolName,

        # all individual stats
        'First' : participant['firstName'],
        'Last' : participant['lastName'],
        'Total_Miles' : round(participant['totalMileage'], 1),
        'Avg_Miles_Per_Day' : round(participant['avgDailyMileage'], 1),
        'Best_Miles' : round(Best_Miles, 2),
        'Best_Miles_Day' : Best_Miles_Date,
        'Total_Steps' : participant['totalSteps'],
        'Avg_Steps_Per_Day' : round(participant['avgDailySteps'], 1),
        'Miles_12_Months' : int(participant['totalMileage']*12),
        'Individual_Plot_Image' : Individual_Chart_Image,

        # team stats
        'Team_Name' : teamName,
        'Team_Total_Miles' : int(df['totalMileage']),
        'Team_Avg_Miles_Per_Day' : round(df['avgMileagePerDay'], 1),
        'Team_Total_Steps' : df['totalSteps'],
        'Team_Avg_Steps_Per_Day' : round(df['avgStepsPerDay'], 1),
        'Team_Plot_Image' : Team_Chart_Image,
        }

    # This is the save version of the document after the context
    folder_name = f'reports/{schoolName}_{teamName}'
    try:
        # Create a new folder in the current working directory
        os.makedirs(folder_name)
        print("Folder created successfully!")
    except FileExistsError:
        # Handle the case where the folder already exists
        print(f"Folder '{folder_name}' already exists.")
    except PermissionError:
        # Handle the case where there are permission issues
        print(f"Permission denied to create folder '{folder_name}'.")
    except Exception as e:
        # Handle any other unexpected errors
        print(f"An error occurred: {e}")

    doc.render(context)
    doc.save(f'reports/{schoolName}_{teamName}/{participant['lastName']}_{participant['firstName']}_report.docx')
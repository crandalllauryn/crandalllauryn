# Set up
import pandas as pd
from datetime import datetime
from matplotlib import pyplot as plt
from pprint import pprint
import re

# datetime object containing current date and time
now = datetime.now()
#today = now.strftime("%Y-%m-%d %H:%M:%S")
print("Today's date and time:", now)

plt.style.use('seaborn-whitegrid')

path = r'C:\Users\crand\OneDrive\Documents\Mission_2.xlsx'

# Bucket names
queue = 'Queue'
examples = 'Examples'
canceled = 'Canceled'
review = 'Review'
in_work = 'In Work'
templates = 'Templates'

emoji_pattern = re.compile("["
                           u"\U0001F600-\U0001F64F"  # emoticons
                           u"\U0001F300-\U0001F5FF"  # symbols & pictographs
                           u"\U0001F680-\U0001F6FF"  # transport & map symbols
                           u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                           "]+", flags=re.UNICODE)

clean_task_name = []

# Pull Teams Excel Data
df = pd.read_excel(path, index_col=0, header=0)
df2 = df.fillna(value=0)

task_name = list(df2.iloc[:, 0])

for i in range(len(task_name)):
    clean_task_name.append(emoji_pattern.sub(r'', task_name[i]))

bucket_data = list(df2.iloc[:, 1])
pull_created_date = list(df2.iloc[:, 6])
pull_completed_date = list(df2.iloc[:, 10])

# Pre-define your dictionaries
scatter_plot_dic = {}
avg_dic = {}
table_dic = {}

# Loop to organize data
for i in range(len(task_name)):
    if pull_completed_date[i] == 0:
        pull_completed_date[i] = pull_created_date[i]
    a = datetime.strptime(pull_created_date[i], '%m/%d/%Y')
    b = datetime.strptime(pull_completed_date[i], '%m/%d/%Y')
    cycle_time = b-a
    wip = now - a
    days = cycle_time.days
    scatter_plot_dic[b] = days
    avg_dic[i] = days
    table_dic[clean_task_name[i]] = [bucket_data[i], days]

pprint(table_dic)
# Find Average Cycle Time
avg = 0
for val in avg_dic.values():
    avg += val

avg = avg // len(avg_dic)

print("Our cycle time average is: " + str(avg) + " days")

# Graph Cycle Time
ax = plt.subplot()
for k, v in scatter_plot_dic.items():
    ax.scatter(k, v,)

ax.axhline(avg, label='Average Cycle Time', linestyle='--')

cell_text = [["1", "1", "1", "1", "1"], ["2", "2", "2", "2", "2"]]
plt.table(cellText=cell_text, loc='bottom')
plt.legend(loc='upper left')
plt.title('Lean Cycle Time')
plt.xlabel('Dates')
plt.ylabel('Cycle Time')
plt.savefig('lean_data.png')

plt.show()

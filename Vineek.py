import pandas as pd
# pd.set_option('display.max_columns', None)
from tabulate import tabulate
from pathlib import Path
from random import random, choice

DIR = Path(__file__).parent
BLINK = 0.182 # random noise

def gen_emptyDatabase():
    times = ['9:30 - 10:30',
             '10:30 - 11:30',
             '11:30 - 12:30',
             '1:30 - 2:30',
             '2:30 - 3:30',
             '3:30 - 4:30',
             '4:30 - 5:30']

    index = pd.MultiIndex.from_product([times,
                                       ['Subject', 'Teacher', 'Room']], names=['Time', 'Details'])
    df = pd.DataFrame(data='', index=index, columns=['Mon', 'Tue', 'Wed', 'Thurs', 'Fri'])
    return df

classes = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Rooms')
main_data = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Timetable')
timetables = dict()


def get_class(amount):
    return choice(classes.loc[classes['Capacity'] <= amount].Class.values)

def checkClash(meeting, day, room, teacher):
    if len(timetables) == 0:
        return False

    previousData = timetables.values()
    for prevData in previousData:
        if (prevData.loc[(meeting, 'Room'), day] == room) or (prevData.loc[(meeting, 'Teacher'), day] == teacher):
            return True
        
    return False

for course_sem, data in main_data.groupby(by=['Course', 'Semester']):
    df = gen_emptyDatabase() # generate timetable to fill
    while not (data['Amount'] == 0).all():
        for meeting in set([meeting[0] for meeting in df.index]):
            for day in df.columns:
                if (df.loc[(meeting, 'Room'), day] == '') and (random() >= BLINK):
                    if len(samples := data.loc[data.Amount > 0]) < 1: # if all subjects have been assigned appropriately
                        break

                    # take a random subject to assign to an empty slot which meets the rules
                    random_subject = samples.sample()
                    capacity = random_subject['Nstudents'].values[0]
                    teacher = random_subject['Teacher'].values[0]
                    room = get_class(capacity)

                    if not checkClash(meeting, day, room, teacher):
                        df.loc[(meeting, 'Room'), day] = room
                        df.loc[(meeting, 'Teacher'), day] = random_subject['Teacher'].values[0]
                        df.loc[(meeting, 'Subject'), day] = random_subject['Subject'].values[0]

                        data.loc[data.Subject == random_subject['Subject'].values[0], 'Amount'] -= 1
    timetables[', '.join([str(x) for x in course_sem])] = df

for course_sem, timetable in timetables.items():
    print(course_sem)
    print(timetable)
    print("#####################################################################################################################################")
    # print(tabulate(timetable.reset_index(level=1), headers='keys', tablefmt='fancy_grid'))
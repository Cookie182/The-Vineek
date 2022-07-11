import pandas as pd
from pathlib import Path
from random import random, choice
from os import system

DIR = Path(__file__).parent
BLINK = 0.1 #! random noise, make sure this is set to a very low float between 0 and 1
TIMES = ['9:30 - 10:30',
         '10:30 - 11:30',
         '11:30 - 12:30',
         '1:30 - 2:30',
         '2:30 - 3:30',
         '3:30 - 4:30',
         '4:30 - 5:30']
NULLVALUE = ''

def loader():
    while True:
        for sign in '_\\|/':
            yield sign

def gen_emptyDatabase():
    index = pd.MultiIndex.from_product([TIMES,
                                       ['Subject', 'Teacher', 'Room']], names=['Time', 'Details'])
    df = pd.DataFrame(data=NULLVALUE, index=index, columns=['Mon', 'Tue', 'Wed', 'Thurs', 'Fri'])
    return df

classes = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Rooms')
main_data = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Timetable')
timetables = dict()


def get_class(amount):
    return choice(classes.loc[classes['Capacity'] <= amount].Class.values)


def checkClash(subject, meeting, day, room, teacher):
    afternoon_times = TIMES[3:]

    if len(timetables) == 0:
        if (('Tut' in subject) or ('Lab' in subject)) and (meeting in afternoon_times):
            return True

        if (('Tut' not in subject) or ('Lab' not in subject)) and (meeting not in afternoon_times):
            return True

    previousData = timetables.values()
    for prevData in previousData:
        # lab/tut in the afternoon rule
        if (('Tut' in subject) or ('Lab' in subject)) and (meeting in afternoon_times):
            if (prevData.loc[(meeting, 'Room'), day] == room) or (prevData.loc[(meeting, 'Teacher'), day] == teacher):
                return False
        else:
            # if random() >= BLINK:
            if (prevData.loc[(meeting, 'Room'), day] == room) or (prevData.loc[(meeting, 'Teacher'), day] == teacher):
                return False

    return True

loader = loader() # funssies

for course_sem, data in main_data.groupby(by=['Course', 'Semester']):
    df = gen_emptyDatabase() # generate timetable to fill
    while not (data['Amount'] == 0).all():
        for meeting in set([meeting[0] for meeting in df.index]):
            for day in df.columns:
                if len(samples := data.loc[data.Amount > 0]) < 1: # if all subjects have been assigned appropriately
                    break

                if (df.loc[(meeting, 'Room'), day] == NULLVALUE) and (random() >= BLINK):
                    # take a random subject to assign to an empty slot which meets the rules
                    random_subject_sample = samples.sample()
                    subject = random_subject_sample['Subject'].values[0]
                    capacity = random_subject_sample['Nstudents'].values[0]
                    teacher = random_subject_sample['Teacher'].values[0]
                    room = get_class(capacity)

                    if checkClash(subject, meeting, day, room, teacher):
                        df.loc[(meeting, 'Room'), day] = room
                        df.loc[(meeting, 'Teacher'), day] = teacher
                        df.loc[(meeting, 'Subject'), day] = subject

                        data.loc[data.Subject == subject, 'Amount'] -= 1

                print(f"Designing timetables ... {next(loader)}", end='\r')

    timetables[', '.join([str(x) for x in course_sem])] = df

system('cls')
for course_sem, timetable in timetables.items():
    print(course_sem)
    print(timetable)
    print("###################################################################################################################################################")
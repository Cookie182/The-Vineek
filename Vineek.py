import pandas as pd
pd.set_option('display.max_columns', 500)
from pathlib import Path
from random import random, choice
from itertools import product
from dataclasses import dataclass
from os import system

DIR = Path(__file__).parent
BLINK = 0.182 #! the higher the value, the more the noise
TIMES = ['9:30 - 10:30',
         '10:30 - 11:30',
         '11:30 - 12:30',
         '1:30 - 2:30',
         '2:30 - 3:30',
         '3:30 - 4:30',
         '4:30 - 5:30']
AFTERNOON_TIMES = TIMES[3:]
DAYS = ['Mon', 'Tue', 'Wed', 'Thurs', 'Fri']
NULLVALUE = ''

@dataclass(frozen=True)
class Vineek:
    DIR = DIR
    BLINK = BLINK
    TIMES = TIMES
    AFTERNOON_TIMES = AFTERNOON_TIMES
    DAYS = DAYS
    NULLVALUE = NULLVALUE
    TRACKCORE_OEL_NULLVALUE = 0
    SUBJECTTYPES = ['Lecture_hrs', 'Lab_hrs', 'Tut_hrs']
    TIMETABLES = dict()

    #! IMPORTANT #!
    subjectsData = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Timetable')
    classesData = pd.read_excel(DIR / 'Time-table.xlsx', header=0, sheet_name='Rooms')
    #! IMPORTANT #!

    def emptyDatabase(self):
        """Generates empty boilerplate of a timetable to be filled in

        Returns:
            Pandas Dataframe: Meeting details as multi-index on the X axis, days of the week as columns
        """
        index = pd.MultiIndex.from_product([self.TIMES,
                                            ['Subject', 'Teacher', 'Room']],
                                           names=['Time', 'Details'])
        emptyDatabase = pd.DataFrame(data=self.NULLVALUE, index=index, columns=self.DAYS)
        return emptyDatabase


    @staticmethod
    def allClassesSlotted(data, subjectTypes):
        """Check if lectures, labs, and tutorials are left to be alloted for a subject

        Args:
            data (pd.DataFrame): Dataframe of details for a course's semester
            subjectTypes (list[str]): Types of classes for that subject

        Returns:
            bool: if all of lectures, labs, tutorials have been allocated for the subject
        """
        return all([(data[subjectType] == 0).all() for subjectType in subjectTypes])


    @staticmethod
    def getRandomSubject(data, subjectTypes):
        """Get random subject and a type of lecture yet to be assigned

        Args:
            data (pd.DataFrame): Dataframe of details for a course's semester
            subjectTypes (list[str]): Types of classes for that subject

        Returns:
            randomSubject, randomSubjectType: Subject and type of lecture
        """
        subjectsTypesCount = dict()
        for course in data.index:
            subjectsTypesCount[course] = []
            for subjectType in subjectTypes:
                if data.loc[course, subjectType] > 0:
                    subjectsTypesCount[course].append(subjectType)

            if len(subjectsTypesCount[course]) == 0: del subjectsTypesCount[course]

        randomSubject = choice(list(subjectsTypesCount.keys()))
        randomSubjectType = choice(subjectsTypesCount[randomSubject])
        return randomSubject, randomSubjectType

    @staticmethod
    def getTeacher(data, subject, subjectType):
        """returns the approriate teacher for the subject depending on type of lecture

        Args:
            data (pd.DataFrame): Dataframe of details for a course's semester
            subject (str): Subject whose teacher/TA details to refer to
            subjectTypes (list[str]): Types of classes for that subject

        Returns:
            str: Name of T.A/faculty depending if the lecture is or isn't a lab/tut respectively
        """
        if subjectType in ['Lab_hrs', 'Tut_hrs']:
            return data.loc[subject, 'TA']
        else:
            return data.loc[subject, 'Faculty']

    def getClass(self, day, time, subjectType, capacity):
        """Returns list of classes eligible to have the class be conducted in depending on class type and capacity

        Args:
            day (str): Day of the week lecture is taking place
            time (str): Timeslot
            subjectType (str): Whether lecture is a lab session or not
            capacity (int): Capacity of the required classroom

        Returns:
            list[int]: List of classrooms that fit the criteria of lecture type and matching capacity
        """
        classType = 'Computer Lab' if subjectType == 'Lab_hrs' else 'Class'

        allClasses = self.classesData.loc[self.classesData['Type'] == classType]
        allClasses = list(allClasses.loc[allClasses['Capacity'] == capacity]['Room_No'].values)

        if len(self.TIMETABLES) == 0:
            return choice(allClasses)
        else:
            for prevTTData in self.TIMETABLES.values():
                occupiedRoom = prevTTData.loc[(time, 'Room'), day]
                if occupiedRoom in allClasses:
                    # print(f"ROOM INFO -> {day}, {time}, {subjectType}, {capacity}")
                    allClasses.remove(prevTTData.loc[(time, 'Room'), day])
        return choice(allClasses)

    def noConsecutiveLectures(self, timetable, day, time, teacher):
        """Prevent consecutive lectures right after another of the same subject and with the same teacher

        Args:
            semesterData (pd.DataFrame): Dataframe of details for a course's semester
            day (str): Day of the week
            time (str): Timeslot
            teacher (str): Name of the teacher/T.A

        Returns:
            bool: Whether the same teacher was found in the previous timeslot
        """

        # if this is the first lecture of the day
        if time == self.TIMES[0]:
            return True

        previousTime = self.TIMES[self.TIMES.index(time)-1]
        if type(teacher) == str: teacher = teacher.split(',')

        previousTeacher = timetable.loc[(previousTime, 'Teacher'), day]
        if type(previousTeacher) == str: previousTeacher.split(',')
            
        for t in teacher:
            # input(t)
            if t in previousTeacher:
                return False

        return True

    def appropriateTime(self, time, labtut):
        if labtut:
            if time in self.AFTERNOON_TIMES:
                return False if random() >= self.BLINK else True
            else:
                return True if random() >= self.BLINK else False

        else:
            if time in self.AFTERNOON_TIMES:
                return True if random() >= self.BLINK else False
            else:
                return False if random() >= self.BLINK else True

    def noClashesCheck(self, subjectType, day, time, teacher, room):
        """Function that iterates through past timetables to check if there are clashes in timeslots

        Args:
            subjectType (str): Whether lecture is a lab session or not
            day (str): Day of the week lecture is taking place
            time (str): Timeslot
            teacher (str): Name of the teacher/T.A
            room (int): Room number to check clashes for

        Returns:
            bool: Returns if there are clashes to be found or not with the current combination on room, day and time
        """
        LabTut = subjectType in ['Lecture_hrs', 'Tut_hrs']

        # if this is the first timetable being created
        if len(self.TIMETABLES) == 0:
            return self.appropriateTime(time, LabTut)

        # incase previous timetables were made
        else:
            if self.appropriateTime(time, LabTut):
                if type(teacher) != list: teacher.split(',')
                if type(room) != list:
                    if type(room) != str:
                        room = str(room).split(',')
                
                for previousTTData in self.TIMETABLES.values():
                    previousRoom = previousTTData.loc[(time, 'Room'), day].split(',')
                    previousTeacher = previousTTData.loc[(time, 'Teacher'), day].split(',')

                    # input((previousRoom, previousTeacher))
                    for t in previousTeacher:
                        if t in teacher:
                            return False

                    for r in previousRoom:
                        if r in room:
                            return False
            return True if random() >= self.BLINK else False


    def getTrackCoreDetails(self, trackCore_data, randomSubjectType, day, time):
        teachers, subjects, classNos = [], [], []
        for subject in trackCore_data.index:
            teachers.append(self.getTeacher(trackCore_data, subject, randomSubjectType))
            subjects.append(subject)

            capacity = trackCore_data.loc[subject, 'Lab_Capacity' if randomSubjectType == 'Lab_hrs' else 'Capacity']
            # input((day, time, trackCore_data, randomSubjectType, capacity))
            classNo = self.getClass(day, time, randomSubjectType, capacity)
            classNos.append(classNo)

        return teachers, subjects, classNos

    def main(self):
        """Fun part"""
        writer = pd.ExcelWriter(self.DIR / "Timetables.xlsx")
        
        for courseSem, semesterData in self.subjectsData.groupby(by=['Dept_id', 'Semester']):
            timetable = self.emptyDatabase()
            semesterData.set_index('Course_Name', inplace=True)
            semesterData['Track_Core'].fillna(self.TRACKCORE_OEL_NULLVALUE, inplace=True)

            while not self.allClassesSlotted(semesterData, self.SUBJECTTYPES):
                # combination of days and timeslots
                for dayMeeting in product(self.DAYS, self.TIMES):
                    day, time = dayMeeting

                    if self.allClassesSlotted(semesterData, self.SUBJECTTYPES): break

                    semesterData = semesterData.sample(frac=1)
                    # input(semesterData)
                    if timetable.loc[(time, 'Room'), day] == self.NULLVALUE:
                        randomSubject, randomSubjectType = self.getRandomSubject(semesterData, self.SUBJECTTYPES)

                        if semesterData.loc[randomSubject, 'Track_Core'] == self.TRACKCORE_OEL_NULLVALUE: # find empty timeslots
                            capacity = semesterData.loc[randomSubject, 'Capacity' if randomSubjectType != 'Lab_hrs' else 'Lab_Capacity']
                            teacher = self.getTeacher(semesterData, randomSubject, randomSubjectType)
                            room = self.getClass(day, time, randomSubjectType, capacity)

                            if not self.noConsecutiveLectures(timetable, day, time, teacher): 
                                break

                            if self.noClashesCheck(randomSubjectType, day, time, teacher, room):
                                timetable.loc[(time, 'Room'), day] = str(room)
                                timetable.loc[(time, 'Teacher'), day] = teacher

                                if randomSubjectType in self.SUBJECTTYPES[1:]:
                                    timetable.loc[(time, 'Subject'), day] = f"{randomSubject} ({randomSubjectType[:3]})"
                                else:
                                    timetable.loc[(time, 'Subject'), day] = randomSubject

                                # keeping track of how many lectures are alloted to a timeslot and how many more need to be allocated
                                semesterData.loc[randomSubject, randomSubjectType] -=1

                                print('#' * 200)
                                print(timetable)
                                print('#' * 200)

                                print(semesterData.to_markdown())

                        else: # for track core/OEl subjects
                            # finding out which TC/OEL subject is and getting other subjects in same TC/OEL group
                            randomSubject_TrackCore = semesterData.loc[randomSubject, 'Track_Core']
                            
                            trackCore_data = semesterData[semesterData['Track_Core'] == randomSubject_TrackCore]

                            teachers, subjects, classNos = self.getTrackCoreDetails(trackCore_data, randomSubjectType, day, time)
                            # input(teachers)

                            # for teacher, subject, classNo in zip(teachers, subjects, classNos):
                            if not self.noConsecutiveLectures(timetable, day, time, teacher):
                                break

                            # for teacher, classNo in zip(teachers, classNos):
                            if not self.noClashesCheck(randomSubjectType, day, time, teachers, classNos):
                                break

                            timetable.loc[(time, 'Room'), day] = ', '.join([str(classNo) for classNo in classNos])
                            timetable.loc[(time, 'Teacher'), day] = ', '.join(teachers)

                            if randomSubjectType in self.SUBJECTTYPES[1:]:
                                timetable.loc[(time, 'Subject'), day] = f"{semesterData.loc[randomSubject, 'Track_Core']} ({randomSubjectType[:3]}) - {', '.join(subjects)}"
                            else:
                                timetable.loc[(time, 'Subject'), day] = f"{semesterData.loc[randomSubject, 'Track_Core']} - {', '.join(subjects)}"

                            for subject in subjects:
                                semesterData.loc[subject, randomSubjectType] -= 1

                            print('#' * 200)
                            print(timetable)
                            print('#' * 200)

                            print(semesterData.to_markdown())

            # saving created semester timetables for each department
            timetable.to_excel(writer, sheet_name=f"{courseSem[0]} - {courseSem[1]}", merge_cells=True)
            self.TIMETABLES[', '.join([str(x) for x in [courseSem]])] = timetable
            input(f"\nTimetable created for {courseSem[0]} - Semester {courseSem[1]}, press any button to continue...\n")
            system('cls')
        writer.close()

if __name__ == '__main__':
    Vineek().main()
import pandas as pd
pd.set_option('display.max_columns', 500)
from pathlib import Path
from random import random, choice
from itertools import product
from dataclasses import dataclass
from os import system

@dataclass
class Vineek:
    DIR = Path(__file__).parent
    BLINK = 0.182 #! the higher the value, the more the noise
    TIMES = ['9:30 - 10:30',
             '10:30 - 11:30',
             '11:30 - 12:30',
             '1:30 - 2:30',
             '2:30 - 3:30',
             '3:30 - 4:30',
             '4:30 - 5:30']
    LECTUT = TIMES[:-2]
    LABTIMES = TIMES[-2:]
    DAYS = ['Mon', 'Tue', 'Wed', 'Thurs', 'Fri']
    NULLVALUE = ''
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

        subjectsTypesCount = []
        for course in data.index:
            for subjectType in subjectTypes:
                if data.loc[course, subjectType] > 0:
                    subjectsTypesCount.append((course, subjectType))

        randomSubject, randomSubjectType = choice(subjectsTypesCount)
        return randomSubject, randomSubjectType

    @staticmethod
    def getTeacher(data, subject, subjectType):
        """returns the approriate teacher for the subject depending on type of lecture

        Args:
            data (pd.DataFrame): Dataframe of details for a course's semester
            subject (str): Subject whose teacher/T.A details to refer to
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
            classData (pd.DataFrame): Dataframe of details for a course's semester
            subjectType (str): Whether lecture is a lab session or not
            capacity (int): Capacity of the required classroom

        Returns:
            list[int]: List of classrooms that fit the criteria of lecture type and matching capacity
        """
        classType = 'Lab' if subjectType == 'Lab_hrs' else 'Class'
        allClasses = self.classesData.loc[self.classesData['Type'] == classType] # getting classroom type
        # input(allClasses)

        allClasses = [str(c) for c in allClasses.loc[allClasses['Capacity'] == capacity]['Room_No'].values]
        # input(allClasses)

        if len(self.TIMETABLES) == 0:
            # input(allClasses)
            # input((classType, subjectType, capacity))
            return choice(allClasses)
        else:
            for prevTTData in self.TIMETABLES.values():
                occupiedRoom = prevTTData.loc[(time, 'Room'), day].split(',')
                allClasses = [aC for aC in allClasses if aC not in occupiedRoom]

        return choice(allClasses)


    def noConsecutiveLectures(self, timetable, day, time, teacher):
        """Prevent consecutive lectures right after another of the same subject and with the same teacher

        Args:
            timetable (pd.DataFrame): Dataframe that is storing the semester's timetable which is being built
            day (str): Day of the week
            time (str): Timeslot
            teacher (str): Name of the teacher/T.A

        Returns:
            bool: Whether the same teacher was found in the previous timeslot
        """

        # if this is the first lecture of the day
        if time == self.LECTUT[0]:
            return True

        previousTime = self.LECTUT[self.LECTUT.index(time)-1]
        if timetable.loc[(previousTime, 'Teacher'), day] == teacher: # check if the previous lecture is the same
            return False

        # if this is not the last lecture of the day
        if time != self.LECTUT[-1]:
            nextTime = self.LECTUT[self.LECTUT.index(time)+1]
            if timetable.loc[(nextTime, 'Teacher'), day] == teacher: # check if the next lecture is the same
                return False

        return True

    def appropriateTime(self, subjectType, time):
        """Function to determine the appropriate timeslot for the lecture depending on the type of lecture it is. Generally returns True if the timeslot for lab sessions are the last 2 timeslots of the day

        Args:
            subjectType (str): Whether lecture is a lab session or not
            time (str): Timeslot

        Returns:
            bool: if timeslot is appropriate for type of lecture
        """
        if subjectType == 'Lab_hrs':
            if time in self.LABTIMES:
                return True
        else:
            if time not in self.LABTIMES:
                return True

        return True if random() < self.BLINK else False

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

        # if this is the first timetable being created
        if len(self.TIMETABLES) == 0:
            return self.appropriateTime(subjectType, time)

        else: # incase previous timetables were made
            if self.appropriateTime(subjectType, time):
                for previousTTData in self.TIMETABLES.values():
                    if (previousTTData.loc[(time, 'Room'), day] == room) or (previousTTData.loc[(time, 'Teacher'), day] == teacher):
                        return False
                return True

    def getTrackCoreDetails(self, trackCore_data, randomSubjectType, day, time):
        """Function that returns the details of all the track cores grouped up in a semester that contains track cores or open electives

        Args:
            trackCore_data (pd.DataFrame): Sub-dataset of the track-core/open elective
            randomSubjectType (str): Whether the lecture is a lab session or not
            day (str): Day of the week the lecture is taking place
            time (str): Timeslot

        Returns:
            tuple(str): Returns the teachers, subjects and the classrooms allocated for the track-core's timeslot
        """
        teachers, subjects, classNos = [], [], []
        for subject in trackCore_data.index:
            teachers.append(self.getTeacher(trackCore_data, subject, randomSubjectType))
            subjects.append(subject)

            capacity = trackCore_data.loc[subject, 'Lab_Capacity' if randomSubjectType == 'Lab_hrs' else 'Capacity']
            # input((day, time, trackCore_data, randomSubjectType, capacity))
            classNo = self.getClass(day, time, randomSubjectType, capacity)
            classNos.append(classNo)

        # input((teachers, subjects, classNos))
        return teachers, subjects, classNos

    def Labs(self, semesterData, timetable):
        """Function that allocates all the lab sessions for a semester. It starts by taking the second last timeslot of an empty timetable and allocating each lab session in succession so each subject has a 2 hour non-stop lab session.

        Args:
            semesterData (pd.DataFrame): Dataframe of the course details for that semester
            timetable (pd.DataFrame): Dataframe that contains the timetable being made for the semester of a course

        Returns:
            semesterData, timetable: Returns updated semester dataframe with updated lab_hrs and timetable with allocated labs
        """
        while not self.allClassesSlotted(semesterData, ['Lab_hrs']):
            firstLabTime = self.LABTIMES[0] # second last time slot
            for day in self.DAYS:
                if self.allClassesSlotted(semesterData, ['Lab_hrs']): break
                randomSubject, randomSubjectType = self.getRandomSubject(semesterData, ['Lab_hrs'])

                if semesterData.loc[randomSubject, 'Track_Core'] == self.TRACKCORE_OEL_NULLVALUE:
                    capacity = semesterData.loc[randomSubject, 'Lab_Capacity']
                    teacher = self.getTeacher(semesterData, randomSubject, randomSubjectType)

                    room = self.getClass(day, firstLabTime, randomSubjectType, capacity)

                    # if not self.noConsecutiveLectures(timetable, day, firstLabTime, teacher): break
                    if self.noClashesCheck(randomSubjectType, day, firstLabTime, teacher, room):
                        for time in self.LABTIMES:
                            timetable.loc[(time, 'Room'), day] = room
                            timetable.loc[(time, 'Teacher'), day] = teacher
                            timetable.loc[(time, 'Subject'), day] = f"{randomSubject} (Lab)"

                        semesterData.loc[randomSubject, 'Lab_hrs'] -= 2

                        print('#' * 200)
                        print(timetable)
                        print('#' * 200)
                        print(semesterData.to_markdown())

                # FOR TRACKCORES/OEL
                else:
                    randomSubject_TrackCore = semesterData.loc[randomSubject, 'Track_Core']
                    trackCore_data = semesterData.loc[semesterData['Track_Core'] == randomSubject_TrackCore]

                    teachers, subjects, classNos = self.getTrackCoreDetails(trackCore_data, 'Lab_hrs', day, firstLabTime)
                    for teacher, subject, classNo in zip(teachers, subjects, classNos):
                        if self.noClashesCheck('Lab_hrs', day, firstLabTime, teacher, classNo):
                            for time in self.LABTIMES:
                                timetable.loc[(time, 'Room'), day] = ', '.join(str(classNo) for classNo in classNos)
                                timetable.loc[(time, 'Teacher'), day] = ', '.join(teachers)
                                timetable.loc[(time, 'Subject'), day] = f"{randomSubject_TrackCore} - {', '.join(subjects)}"

                            semesterData.loc[subject, randomSubjectType] -= 2

                    print('#' * 200)
                    print(timetable)
                    print('#' * 200)

                    print(semesterData.to_markdown())
            if self.allClassesSlotted(semesterData, ['Lab_hrs']): break

        return semesterData, timetable

    def LecturesTuts(self, semesterData, timetable):
        """Same as the labs function, though for lectures and tutorials

        Args:
            semesterData (pd.DataFrame): Dataframe of the course details for that semester
            timetable (pd.DataFrame): Dataframe that contains the timetable being made for the semester of a course

        Returns:
            semesterData, timetable: Returns updated semester dataframe with updated lab_hrs and timetable with allocated lectures and tutorials
        """
        while not self.allClassesSlotted(semesterData, ['Lecture_hrs', 'Tut_hrs']):
            # combination of days and timeslots
            for dayMeeting in product(self.LECTUT, self.DAYS):
                time, day = dayMeeting

                if self.allClassesSlotted(semesterData, ['Lecture_hrs', 'Tut_hrs']): break

                if timetable.loc[(time, 'Room'), day] == self.NULLVALUE:
                    randomSubject, randomSubjectType = self.getRandomSubject(semesterData, ['Lecture_hrs', 'Tut_hrs'])

                    if semesterData.loc[randomSubject, 'Track_Core'] == self.TRACKCORE_OEL_NULLVALUE: # for regular subjects
                        capacity = semesterData.loc[randomSubject, 'Capacity']
                        teacher = self.getTeacher(semesterData, randomSubject, randomSubjectType)
                        room = self.getClass(day, time, randomSubjectType, capacity)

                        if not self.noConsecutiveLectures(timetable, day, time, teacher): break

                        if self.noClashesCheck(randomSubjectType, day, time, teacher, room):
                            timetable.loc[(time, 'Room'), day] = room
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

                        for teacher, subject, classNo in zip(teachers, subjects, classNos):
                            if not self.noConsecutiveLectures(timetable, day, time, teacher):
                                break

                            if not self.noClashesCheck(randomSubjectType, day, time, teacher, classNo):
                                break

                        timetable.loc[(time, 'Room'), day] = ', '.join([str(classNo) for classNo in classNos])
                        timetable.loc[(time, 'Teacher'), day] = ', '.join(teachers)
                        timetable.loc[(time, 'Subject'), day] = f"{semesterData.loc[randomSubject, 'Track_Core']} - {', '.join(subjects)}"

                        for subject in subjects:
                            semesterData.loc[subject, randomSubjectType] -= 1

                        print('#' * 200)
                        print(timetable)
                        print('#' * 200)

                        print(semesterData.to_markdown())

        return semesterData, timetable

    def main(self):
        """Fate fell short this time"""
        for courseSem, semesterData in self.subjectsData.groupby(by=['Dept_id', 'Semester']):
            timetable = self.emptyDatabase()
            semesterData.set_index('Course_Name', inplace=True)
            semesterData['Track_Core'].fillna(self.TRACKCORE_OEL_NULLVALUE, inplace=True)

            # for labs
            semesterData, timetable = self.Labs(semesterData, timetable)

            print('#' * 200)
            print(timetable)
            print('#' * 200)
            print(semesterData.to_markdown())

            # for lectures and tutorials
            semesterData, timetable = self.LecturesTuts(semesterData, timetable)

            # saving created semester timetables for each department
            self.TIMETABLES[', '.join([str(x) for x in [courseSem]])] = timetable
            input(f"\nTimetable created for {courseSem[0]} - Semester {courseSem[1]}, press any button to continue...\n")
            system('cls')

if __name__ == '__main__':
    Vineek().main()
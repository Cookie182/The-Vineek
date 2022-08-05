import pandas as pd
pd.set_option('display.max_columns', 500)
from pathlib import Path
from random import choice
from itertools import product
from os import system
from time import sleep


class Vineek:
    def __init__(self, TIMES, LABTIMES, DAYS, NULLVALUE=''):
        self.DIR = self.makeDir()
        self.TIMES = TIMES
        self.LABTIMES = LABTIMES
        self.DAYS = DAYS
        self.SUBJECTTYPES = ['Lecture_hrs', 'Lab_hrs', 'Tut_hrs']
        self.NULLVALUE = NULLVALUE
        self.fileName = 'Time-table.xlsx'

        if (self.DIR / self.fileName).is_file():
            print(f"File found {self.DIR / self.fileName}")
            self.subjectsData = pd.read_excel(self.DIR / self.fileName, header=0, sheet_name='Timetable')
            self.classesData = pd.read_excel(self.DIR / self.fileName, header=0, sheet_name='Rooms')
            input("Excel found for timetable data found, please make sure everything is entered corerctly before pressing Enter...")
        else:
            self.createExcelFile()
            self.subjectsData = pd.read_excel(self.DIR / self.fileName, header=0, sheet_name='Timetable')
            self.classesData = pd.read_excel(self.DIR / self.fileName, header=0, sheet_name='Rooms')

        self.TIMETABLES = dict()
        self.facultyTT = dict()
        self.roomTT = dict()

        self.main()


    @staticmethod
    def makeDir():
        """Func that finds the path to the desktop of the user to store timetable and store created timetables"""
        path = Path("C:\\Users")
        for folder in path.glob('*'):
            if folder.parts[-1] not in ['All Users', 'Default', 'Default User', 'desktop.ini', 'Public']:
                path_ = path / folder.parts[-1] / 'Desktop'
                if path_.is_dir():
                    path = path / folder.parts[-1] / 'Desktop'
                    break

        print(f"Everything will be stored in {path}\n")
        sleep(3)
        return path


    def createExcelFile(self):
        """Function to create the prerequisite excel files needed for gathering appropriate data
        """
        timetableColumns = ['Dept_id', 'Course_id', 'Track_Core', 'Course_Name', 'Faculty', 'TA', 'Semester', 'Lecture_hrs',
                            'Tut_hrs', 'Capacity', 'Lab_hrs', 'Lab_Capacity', 'Assigned_Room', 'Assigned_Lab']
        roomsColumns = ['Room_No', 'Capacity', 'Type']

        writer = pd.ExcelWriter(self.DIR / self.fileName)

        pd.DataFrame(columns=timetableColumns).set_index(timetableColumns[0]).to_excel(writer, sheet_name='Timetable')
        pd.DataFrame(columns=roomsColumns).set_index(roomsColumns[0]).to_excel(writer, sheet_name='Rooms')
        writer.save()
        input('Excel file created...please input details approriately before continuing ...')
        system('cls')


    def emptyTimetable(self):
        """Generate an empty timetable to input data into for each semester and for each course

        Returns:
            pd.DataFrame: Empty dataframe with appropriate columns and timeslot data
        """
        index = pd.MultiIndex.from_product([self.TIMES,
                                            ['Subject', 'Teacher', 'Room']],
                                           names=['Time', 'Details'])
        emptyDatabase = pd.DataFrame(data=self.NULLVALUE, index=index, columns=self.DAYS)
        return emptyDatabase


    def getLabTimes(self):
        """Function that iterates through the Lab times for unique batch for each course of

        Yields:
            str: timeslot to allocate labs
        """
        while True:
            for time in self.LABTIMES:
                yield time


    @staticmethod
    def allClasssesSlotted(semesterData, subjectTypes):
        """Function to check if all the classes for a semester batch has yet to be  allocated on the timetable

        Args:
            semesterData (pd.DataFrame): Dataframe of semester specific data for a course of
            subjectTypes (pd.DataFrame): list of the type of subjects to retrieve data for each

        Returns:
            _type_: _description_
        """
        for subjectType in subjectTypes:
            for _, row in semesterData.iterrows():
                if row[subjectType] > 0:
                    return False

        return True


    def getRandomSubject(self, semesterData, subjectTypes):
        """Func that returns a random subject and type of lecture that is yet to be alloted

        Args:
            semesterData (pd.DataFrame): Pandas Dataframe of the timetable which contains data for that specific semesterData
            subjectTypes (list(string)): list of strings that contain the lecture types to be assigned for that subject

        Returns:
            string, string: returns a random subject and the type of lecture of that subject to be assigned
        """
        subjectTypesCount = []
        for subject in semesterData.index:
            for subjectType in subjectTypes:
                if semesterData.loc[subject, subjectType] > 0:
                    subjectTypesCount.append((subject, subjectType))

        randomSubject, randomSubjectType = choice(subjectTypesCount)
        return randomSubject, randomSubjectType
    

    @staticmethod
    def getTeacher(semesterData, subject, subjectType):
        """Function to get the teacher/TA for that subject and subject type of

        Args:
            semesterData (pd.DataFrame): Pandas Dataframe of the timetable which contains data for that specific semesterData
            subject (str): Name of the subject
            subjectType (str): Type of lecture for that subject. Can be normal lecture, tutorial or lab

        Returns:
            str: Name of the teacher/TA depending on subject type
        """
        return semesterData.loc[subject, 'TA' if subjectType in ['Lab_hrs', 'Tut_hrs'] else 'Faculty']


    def getClass(self, semesterData, subject, day, time, subjectType, capacity):
        """Function to get a random classroom for the lecture to take place in or to retrieve the assigned room/lab for that subject

        Args:
            semesterData (pd.DataFrame): Pandas Dataframe of the timetable which contains data for that specific semesterData
            subject (str): Nameo f the subject
            day (str): Day of the week
            time (str): Timeslot
            subjectType (str): Type of lecture for that subject. Can be normal lecture, tutorial or lab
            capacity (int): Capacity of the classroom/lab that the lecture needs

        Returns:
            int: returns appropriate room/lab number
        """

        classType = 'Lab' if subjectType == 'Lab_hrs' else 'Class'
        assignedRoomLabel = f"Assigned_{'Lab' if subjectType == 'Lab_hrs' else 'Room'}"

        if semesterData.loc[subject, assignedRoomLabel] == self.NULLVALUE:
            """FOR SUBJECTS WITH NO RESERVED ROOMS"""
            allClasses = self.classesData[(self.classesData['Type'] == classType) & (self.classesData['Capacity'] == capacity)]['Room_No']
            allClasses = [str(int(class_)) for class_ in allClasses]

            if len(self.TIMETABLES) == 0:
                return choice(allClasses)
            else:
                for prevTTData in self.TIMETABLES.values():
                    occupiedRooms = prevTTData.loc[(time, 'Room'), day].split(',')
                    allClasses = [class_ for class_ in allClasses if class_ not in occupiedRooms]

                return choice(allClasses)

        else:
            """FOR SUBJECTS WITH RESERVED ROOMS"""
            assignedRoom = str(semesterData.loc[subject, assignedRoomLabel])
            if len(self.TIMETABLES) == 0:
                return assignedRoom
            else:
                for prevTTData in self.TIMETABLES.values():
                    occupiedRooms = prevTTData.loc[(time, 'Room'), day].split(',')
                    if assignedRoom not in occupiedRooms:
                        return assignedRoom

            return None


    def noClashesCheck(self, day, time, teacher, room):
        """Function to check if there are classes in that timeslot for that specific day by checking the room and teachers with previously made timetables

        Args:
            day (str): Day of the week
            time (str): Timeslot
            teacher (str): Name of the teacher/TA
            room (int): room no.

        Returns:
            bool: returns if there are clashes found or not
        """
        if len(self.TIMETABLES) == 0:
            return True

        else:
            for previousTTData in self.TIMETABLES.values():
                if (previousTTData.loc[(time, 'Room'), day] == room) or (previousTTData.loc[(time, 'Teacher'), day] == teacher):
                    return False

            return True

    def assignRoomFacultyTT(self, facultyName, day, time, subjectName, room):
        """Function to create and allocate details into the faculty timetable for that specific teacher/TA and room number

        Args:
            facultyName (str): Name of teacher/TA
            day (str): Day of the week
            time (str): Timeslot
            subjectName (str): Name of the subjectName
            room (int): room/lab number
        """
        if facultyName not in self.facultyTT.keys():
            self.facultyTT[facultyName] = self.emptyTimetable()

        if room not in self.roomTT.keys():
            self.roomTT[room] = self.emptyTimetable()

        # faculty timetable
        self.facultyTT[facultyName][(time, 'Room'), day] = room
        self.facultyTT[facultyName][(time, 'Subject'), day] = subjectName
        self.facultyTT[facultyName][(time, 'Teacher'), day]  = facultyName

        # room timetable
        self.roomTT[room][(time, 'Room'), day] = room
        self.roomTT[room][(time, 'Subject'), day] = subjectName
        self.roomTT[room][(time, 'Teacher'), day]  = facultyName


    def Labs(self, semesterData, timetable, labTime):
        """Function that allocates the labs for

        Args:
            semesterData (_type_): _description_
            timetable (_type_): _description_
            labTime (_type_): _description_

        Returns:
            _type_: _description_
        """
        semesterLabData = semesterData[semesterData['Lab_hrs'] > 0]
        consecutiveLabTime = self.TIMES[self.TIMES.index(labTime) + 1]

        while not self.allClasssesSlotted(semesterLabData, ['Lab_hrs']):
            for day in self.DAYS:
                if self.allClasssesSlotted(semesterLabData, ['Lab_hrs']) == True:
                    break

                while True:
                    randomSubject, randomSubjectType = self.getRandomSubject(semesterLabData, ['Lab_hrs'])

                    if semesterLabData.loc[randomSubject, 'Track_Core'] == self.NULLVALUE:
                        capacity = semesterLabData.loc[randomSubject, 'Lab_Capacity']
                        teacher = self.getTeacher(semesterData, randomSubject, randomSubjectType)

                        room = self.getClass(semesterLabData, randomSubject, day, labTime, randomSubjectType, capacity)
                        if room == None:
                            continue

                        if self.noClashesCheck(day, labTime, teacher, room):
                            for time in [labTime, consecutiveLabTime]:
                                timetable.loc[(time, 'Room'), day] = room
                                timetable.loc[(time, 'Teacher'), day] = teacher
                                timetable.loc[(time, 'Subject'), day] = f"{randomSubject} (Lab)"

                                self.assignRoomFacultyTT(teacher, day, time, f"{randomSubject} (Lab)", room)

                            semesterData.loc[randomSubject, randomSubjectType] -= 2
                            semesterLabData.loc[randomSubject, randomSubjectType] -= 2

                            print('#' * 200)
                            print(timetable)
                            print('#' * 200)
                            print(semesterData.to_markdown())
                            break

                    else:
                        randomSubject_TrackCore = semesterLabData.loc[randomSubject, 'Track_Core']
                        trackCoreData = semesterLabData.loc[semesterLabData['Track_Core'] == randomSubject_TrackCore]

                        subjects = trackCoreData.index
                        teachers = [self.getTeacher(trackCoreData, subject, randomSubjectType) for subject in subjects]
                        classNos = []

                        capacities = trackCoreData.loc[subjects, 'Lab_Capacity' if randomSubjectType == 'Lab_hrs' else 'Capacity']
                        for subject, capacity in zip(subjects, capacities):
                            while True:
                                classNo = self.getClass(trackCoreData, subject, day, labTime, randomSubjectType, capacity)
                                if classNo not in classNos:
                                    classNos.append(classNo)
                                    break

                        if None in classNos:
                            """No approriate class found in this timeslot"""
                            continue

                        clashFound = False
                        for teacher, subject, classNo in zip(teachers, subjects, classNos):
                            if self.noClashesCheck(day, labTime, teacher, classNo) == False:
                                clashFound = True
                                break

                        if clashFound == True:
                            continue

                        for teacher, subject, classNo in zip(teachers, subjects, classNos):
                            for time in [labTime, consecutiveLabTime]:
                                timetable.loc[(time, 'Room'), day] = ', '.join(classNos)
                                timetable.loc[(time, 'Teacher'), day] = ', '.join(teachers)
                                trackCoreLabName = f"{randomSubject_TrackCore} (Lab) - {', '.join(subjects)}"
                                timetable.loc[(time, 'Subject'), day] = trackCoreLabName

                                self.assignRoomFacultyTT(teacher, day, time, f"{randomSubject_TrackCore} (Lab) - {subject}", classNo)

                            semesterData.loc[subject, randomSubjectType] -= 2
                            semesterLabData.loc[subject, randomSubjectType] -= 2

                        print('#' * 200)
                        print(timetable)
                        print('#' * 200)
                        print(semesterData.to_markdown())
                        break

            if self.allClasssesSlotted(semesterLabData, ['Lab_hrs']): break

        return (semesterData, timetable)
    

    def noConsecutiveLectures(self, timetable, day, time, teacher):
        if time == self.TIMES[0]:
            return True

        previousTime = self.TIMES[self.TIMES.index(time)-1]
        if timetable.loc[(previousTime, 'Teacher'), day] == teacher: # check if the previous lecture is the same
            return False

        # if this is not the last lecture of the day
        if time != self.TIMES[-1]:
            nextTime = self.TIMES[self.TIMES.index(time)+1]
            if timetable.loc[(nextTime, 'Teacher'), day] == teacher: # check if the next lecture is the same
                return False

        return True


    def LecturesTuts(self, semesterData, timetable):
        while not self.allClasssesSlotted(semesterData, ['Lecture_hrs', 'Tut_hrs']):
            for time, day in product(self.TIMES, self.DAYS):
                if self.allClasssesSlotted(semesterData, ['Lecture_hrs', 'Tut_hrs']): break

                if timetable.loc[(time, 'Room'), day] == self.NULLVALUE:
                    for _ in range(len(semesterData)):
                        randomSubject, randomSubjectType = self.getRandomSubject(semesterData, ['Lecture_hrs', 'Tut_hrs'])
                        notTrackcore = all(x == self.NULLVALUE for x in semesterData.loc[:, 'Track_Core'])

                        if notTrackcore == True:
                            capacity = semesterData.loc[randomSubject, 'Capacity']
                            teacher = self.getTeacher(semesterData, randomSubject, randomSubjectType)
                            room = self.getClass(semesterData, randomSubject, day, time, randomSubjectType, capacity)

                            if room is None:
                                continue

                            if not self.noConsecutiveLectures(timetable, day, time, teacher): continue

                            if self.noClashesCheck(day, time, teacher, room):
                                timetable.loc[(time, 'Room'), day] = room
                                timetable.loc[(time, 'Teacher'), day] = teacher

                                subjectName = randomSubject if randomSubjectType != 'Tut_hrs' else f"{randomSubject} (Tut)"
                                timetable.loc[(time, 'Subject'), day] = subjectName

                                self.assignRoomFacultyTT(teacher, day, time, subjectName, room)
                                semesterData.loc[randomSubject, randomSubjectType] -=1

                                print('#' * 200)
                                print(timetable)
                                print('#' * 200)

                                print(semesterData.to_markdown())
                                break

                        else:
                            """FOR TRACKCORES/OELS"""
                            randomSubject_TrackCore = semesterData.loc[randomSubject, 'Track_Core']
                            trackCoreData = semesterData[semesterData['Track_Core'] == randomSubject_TrackCore]

                            teachers, subjects, classNos = [], [], []
                            for subject in trackCoreData.index:
                                if subject not in subjects:
                                    teacher = self.getTeacher(trackCoreData, subject, randomSubjectType)
                                    if (teacher not in teachers) and (trackCoreData.loc[subject, randomSubjectType] > 0):
                                        teachers.append(teacher)
                                        subjects.append(subject)

                                        capacity = trackCoreData.loc[subject, 'Lab_Capacity' if randomSubjectType == 'Lab_hrs' else 'Capacity']
                                        classNo = self.getClass(trackCoreData, subject, day, time, randomSubjectType, capacity)
                                        classNos.append(classNo)

                            if None in classNos:
                                continue
                            classNos = [str(int(float(classNo))) for classNo in classNos]

                            noconsecutiveLectures = True
                            noClashesCheck = True
                            for teacher, subject, classNo in zip(teachers, subjects, classNos):
                                if not self.noConsecutiveLectures(timetable, day, time, teacher):
                                    noconsecutiveLectures = False

                                if not self.noClashesCheck(day, time, teacher, classNo):
                                    noClashesCheck = False

                            if (not noconsecutiveLectures) or (not noClashesCheck):
                                continue

                            for subject in subjects:
                                semesterData.loc[subject, randomSubjectType] -= 1

                            timetable.loc[(time, 'Room'), day] = ', '.join([str(classNo) for classNo in classNos])
                            timetable.loc[(time, 'Teacher'), day] = ', '.join(teachers)

                            if randomSubjectType == 'Tut_hrs':
                                timetable.loc[(time, 'Subject'), day] = f"{semesterData.loc[randomSubject, 'Track_Core']} (Tut) - {', '.join(subjects)}"
                            else:
                                timetable.loc[(time, 'Subject'), day] = f"{semesterData.loc[randomSubject, 'Track_Core']} - {', '.join(subjects)}"

                            for teacher, subject, classNo in zip(teachers, subjects, classNos):
                                self.assignRoomFacultyTT(teacher, day, time, f"{randomSubject_TrackCore} - {subject} {'(Tut)' if randomSubjectType == 'Tut_hrs' else ''}", classNo)

                            print('#' * 200)
                            print(timetable)
                            print('#' * 200)

                            print(semesterData.to_markdown())
                            break

        return semesterData, timetable


    def saveTables(self):
        ttPath = self.DIR / 'Vineek Timetables'
        ttPath.mkdir(parents=True, exist_ok=True)

        roomPath = self.DIR / 'Vineek Room Timetables'
        roomPath.mkdir(parents=True, exist_ok=True)

        facultyPath = self.DIR / 'Vineek Faculty Timetables'
        facultyPath.mkdir(parents=True, exist_ok=True)

        # saving lecture, room and faculty timetables
        for ttName, TT in self.TIMETABLES.items():
            TT.to_excel(ttPath / f"{ttName}.xlsx", merge_cells=True)
        print(f"\nAll lecture timetables have been saved in {ttPath}\n")

        for ttName, TT in self.roomTT.items():
            TT.to_excel(roomPath / f"{ttName}.xlsx", merge_cells=True)
        print(f"\nAll room timetables have been saved in {roomPath}\n")

        for ttName, TT in self.facultyTT.items():
            TT.to_excel(facultyPath / f"{ttName}.xlsx", merge_cells=True)
        print(f"\nAll faculty timetables have been saved in {facultyPath}\n")


    def main(self):
        labTime = self.getLabTimes()

        for courseSem, semesterDataMain in self.subjectsData.groupby(by=['Dept_id', 'Semester']):
            semesterDataMain.set_index('Course_Name', inplace=True)
            semesterDataMain.fillna(self.NULLVALUE, inplace=True)

            batchCount = int(input(f"How many batches are there for {courseSem[0]} - Semester {courseSem[1]}? "))
            for _ in range(batchCount):
                timetable = self.emptyTimetable()

                """LABS"""
                semesterData, timetable = self.Labs(semesterDataMain.copy(), timetable, next(labTime))

                print('#' * 200)
                print(timetable)
                print('#' * 200)
                print(semesterData.to_markdown())

                semesterData, timetable = self.LecturesTuts(semesterData, timetable)

                self.TIMETABLES['-'.join([str(x) for x in courseSem])] = timetable

        print("\nTimetables generated, saving them in their respective folders...\n")
        self.saveTables()

        input("\nAll timetables generated!")


TIMES = ['9:30 AM - 10:30 AM',
        '10:30 AM - 11:30 AM',
        '11:30 AM - 12:30 PM',
        '1:30 PM - 2:30 PM',
        '2:30 PM - 3:30 PM',
        '3:30 PM - 4:30 PM',
        '4:30 PM - 5:30 PM']

LABTIMES = ['3:30 PM - 4:30 PM',
            '10:30 AM - 11:30 AM',
            '1:30 PM - 2:30 PM']

DAYS = ['Mon', 'Tue', 'Wed', 'Thurs', 'Fri']

Vineek(TIMES = TIMES,
       LABTIMES = LABTIMES,
       DAYS = DAYS)
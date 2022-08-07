from pandas import DataFrame, MultiIndex, read_excel, ExcelWriter, set_option
set_option('display.max_columns', 500)
from pathlib import Path
from random import choice
from itertools import product
from os import system, getlogin
from time import sleep


class Vineek:
    def __init__(self, TIMESLOTS, DAYS, NULLVALUE=''):
        self.DIR = Path(f"C:\\Users\\{getlogin()}\\Desktop") # get the directory to the current windows user's desktop
        self.TIMESLOTS = TIMESLOTS
        self.LABTIMES = self.generateLabTimes()
        self.DAYS = DAYS
        self.SUBJECTTYPES = ['Lecture_hrs', 'Lab_hrs', 'Tut_hrs']
        self.NULLVALUE = NULLVALUE
        self.fileName = 'Time-table.xlsx'

        # Rules for timetable detail data entry
        print("Welcome to the Vineek timetable generation algorithm! Do note that if it looks like the program is stuck, simply close the window and start it up again!\n\nCredits:\nAshwin Rajesh Jawalikar\nGurvinder Kaur\n\nRules:\n- All the timetables, along with the initial excel sheet to provide the necessary data will be on the desktop. If not, the file directory path for them will be shown\n- 'Lecture_hrs', 'Tut_hrs', 'Lab_hrs' represents the amount of lectures of each respective type of lecture there is for a particular subject.\n- If a subject does not have any lab or tutorial lectures, the 'TA' column for that subject should be left as blank\n- Continuing from the previous point, put a 0 respective columns of 'Lecture_hrs', 'Lab_hrs', 'Tut_hrs' if the subject doesn't have any of those respective type of classes\n- 'Assigned_Room' and 'Assigned_Lab' columns are for pre-allocating specific rooms for a subject, either for a lecture or for a lab/tutorial session respectively\n- Please do not use the same values for 'Course_id' and 'Track_Core' columns\n- Please do make sure all variables in the 'Course_id' feature are unique as they play a integral role for track cores\n- Speaking of track cores, if a subject is common in a track core, allocate them in the same slot, assuming they also have the same teacher.\n- Try to keep the sample size small. To do this, try to only specific specific rooms to use for a specific course as the algorithm may panic and/or freeze\n- All the final outputs are in an excel format for further modifications and/or to cross reference to make more subjective changes\n -If you do wish to change the timeslots for all lectures, or even timeslots for labs as the amount of labs (assuming there are labs for the subject, are in multiples of 2) will be held consecutively, they can be changed within the code itself but in a very easy format just by editing the timeslots respectively near the end of the .py script.\n- Be gentle, she's a shy kind-hearted soul\n")
        
        if (self.DIR / self.fileName).is_file():
            print(f"File found {self.DIR / self.fileName}")
            self.subjectsData = read_excel(self.DIR / self.fileName, header=0, sheet_name='Timetable')
            self.classesData = read_excel(self.DIR / self.fileName, header=0, sheet_name='Rooms')
            input("Excel found for timetable data found, please make sure everything is entered correctly before pressing Enter...")
            system('cls')
        else:
            self.createTTDataExcelFile()

        self.TIMETABLES = dict() # store lecture timetables for each batch, semester and course
        self.facultyTT = dict() # store faculty timetables
        self.roomTT = dict() # store room timetables
        self.main() # absolute war


    def createTTDataExcelFile(self):
        """Function to create the prerequisite excel file needed for gathering appropriate data for timetable creation"""

        timetableColumns = ['Dept_id', 'Course_id', 'Track_Core', 'Course_Name', 'Faculty', 'TA', 'Semester', 'Lecture_hrs',
                            'Tut_hrs', 'Capacity', 'Lab_hrs', 'Lab_Capacity', 'Assigned_Room', 'Assigned_Lab']
        roomsColumns = ['Room_No', 'Capacity', 'Type']

        writer = ExcelWriter(self.DIR / self.fileName)

        DataFrame(columns=timetableColumns).set_index(timetableColumns[0]).to_excel(writer, sheet_name='Timetable')
        DataFrame(columns=roomsColumns).set_index(roomsColumns[0]).to_excel(writer, sheet_name='Rooms')
        writer.save()
        input(f'Excel file created, path: {self.DIR / self.fileName}...please input details approriately to start generating timetables. The program is now going to exit, simply start the program again after entering the necessary timetable data.')
        exit()


    def emptyTimetable(self):
        """Generate an empty timetable to input data into for each semester and for each course

        Returns:
            pd.DataFrame: Empty dataframe with appropriate columns and timeslot data
        """

        index = MultiIndex.from_product([self.TIMESLOTS,
                                            ['Subject', 'Teacher', 'Room']],
                                           names=['Time', 'Details'])
        emptyDatabase = DataFrame(data=self.NULLVALUE, index=index, columns=self.DAYS)
        return emptyDatabase


    def generateLabTimes(self):
        """Generate a list of timeslots that can hold uninterrupted 2 hour lab sessions for subjects

        Returns:
            list: list of timeslots
        """
        
        labTimes = []
        for timeslotIndex, timeslot in enumerate(self.TIMESLOTS[:-1]):
            timeslotEndTime = timeslot.split(' - ')[-1]
            if timeslotEndTime in self.TIMESLOTS[timeslotIndex+1]:
                labTimes.append(timeslot)

        return labTimes
    

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
            bool: Whether or not if there are any more lectures of all types left to allocate for all subjects in a semester
        """

        for subjectType in subjectTypes:
            for _, row in semesterData.iterrows():
                if row[subjectType] > 0:
                    return False

        return True


    @staticmethod
    def getRandomSubject(semesterData, subjectTypes):
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
        """Function to get the teacher/TA for that subject depending on the type of lecture

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
        if (facultyName != '') or (not facultyName.isspace()): # prevent generation of faculty TT with blank names
            self.facultyTT[facultyName].loc[(time, 'Room'), day] = room
            self.facultyTT[facultyName].loc[(time, 'Subject'), day] = subjectName
            self.facultyTT[facultyName].loc[(time, 'Teacher'), day]  = facultyName

        # room timetable
        self.roomTT[room].loc[(time, 'Room'), day] = room
        self.roomTT[room].loc[(time, 'Subject'), day] = subjectName
        self.roomTT[room].loc[(time, 'Teacher'), day]  = facultyName


    def Labs(self, semesterData, timetable, labTime):
        """Function that allocates the labs in the semester timetable from the semester data

        Args:
            semesterData (pd.DataFrame): Pandas Dataframe which contains data for that specific semesterData
            timetable (pd.DataFrame): Pandas Dataframe of the timetable for the semester
            labTime (str): Timeslot for the initial lab lecture to allocate

        Returns:
            (pd.DataFrame, pd.DataFrame): returns the semesterData and timetable
        """

        semesterLabData = semesterData[semesterData['Lab_hrs'] > 0]
        consecutiveLabTime = self.TIMESLOTS[self.TIMESLOTS.index(labTime) + 1]

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
                            print(semesterData)
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
                        print(semesterData)
                        break

            if self.allClasssesSlotted(semesterLabData, ['Lab_hrs']): break

        return (semesterData, timetable)


    def noConsecutiveLectures(self, timetable, day, time, teacher):
        """Check whether the next and previous lecture of the current timeslot contains the same teacher to prevent teachers from having consecutive lectures

        Args:
            timetable (pd.DataFrame): Pandas Dataframe of the semester timetable
            day (str): Day of the weeek
            time (str): Timeslot
            teacher (str): Name of the teacher/TA

        Returns:
            bool: Whether the previous or next lecture from the given timeslot contains the same teacher
        """

        if time == self.TIMESLOTS[0]:
            return True

        previousTime = self.TIMESLOTS[self.TIMESLOTS.index(time)-1]
        if timetable.loc[(previousTime, 'Teacher'), day] == teacher: # check if the previous lecture is the same
            return False

        # if this is not the last lecture of the day
        if time != self.TIMESLOTS[-1]:
            nextTime = self.TIMESLOTS[self.TIMESLOTS.index(time)+1]
            if timetable.loc[(nextTime, 'Teacher'), day] == teacher: # check if the next lecture is the same
                return False

        return True


    def LecturesTuts(self, semesterData, timetable):
        """Function that allocates the lectures/tutorials in the semester timetable from the semester data

        Args:
            semesterData (pd.DataFrame): Pandas Dataframe which contains data for that specific semesterData
            timetable (pd.DataFrame): Pandas Dataframe of the timetable for the semester

        Returns:
            (pd.DataFrame, pd.DataFrame): returns the semesterData and timetable
        """

        while not self.allClasssesSlotted(semesterData, ['Lecture_hrs', 'Tut_hrs']):
            for time, day in product(self.TIMESLOTS, self.DAYS):
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
                                print(semesterData)
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

                            print(semesterData)
                            break

        return semesterData, timetable


    def saveTables(self):
        """Function to save the timetables for all types of lectures, rooms and faculties into an organised excel directory """

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
            if ttName != '':
                TT.to_excel(facultyPath / f"{ttName}.xlsx", merge_cells=True)
        print(f"\nAll faculty timetables have been saved in {facultyPath}\n")


    def main(self):
        labTime = self.getLabTimes()

        for courseSem, semesterDataMain in self.subjectsData.groupby(by=['Dept_id', 'Semester']):
            semesterDataMain.set_index('Course_Name', inplace=True)
            semesterDataMain.fillna(self.NULLVALUE, inplace=True)

            while True:
                try:
                    batchCount = int(input(f"How many batches are there for {courseSem[0]} - Semester {courseSem[1]}? (Maximum can be 4!): "))
                    if batchCount > 4: raise ValueError("Number of batches too high!")
                    break
                except:
                    print("\nPlease enter a valid number of batches!")

            for batchNo in range(batchCount):
                sleep(1)
                timetable = self.emptyTimetable()

                """LABS"""
                semesterData, timetable = self.Labs(semesterDataMain.copy(), timetable, next(labTime))

                print('#' * 200)
                print(timetable)
                print('#' * 200)
                print(semesterData)

                """FOR LECTURES AND TUTORIALS"""
                semesterData, timetable = self.LecturesTuts(semesterData, timetable)

                """SAVING THE TIMETABLE"""
                timetableName = f"{courseSem[0]} - Semester {courseSem[1]}"
                if batchCount < 2:
                    self.TIMETABLES[timetableName] = timetable
                else:
                    timetableName += f" - Batch {batchNo+1}"
                    self.TIMETABLES[timetableName] = timetable

        print("\nTimetables generated, saving them in their respective folders...\n")
        sleep(1)
        self.saveTables() # creating room & faculty timetables then saving them on top of the timetables for lectures

        input("\nAll timetables generated!")


TIMESLOTS = ['9:30 AM - 10:30 AM',
             '10:30 AM - 11:30 AM',
             '11:30 AM - 12:30 PM',
             '1:30 PM - 2:30 PM',
             '2:30 PM - 3:30 PM',
             '3:30 PM - 4:30 PM',
             '4:30 PM - 5:30 PM']

DAYS = ['Mon', 'Tue', 'Wed', 'Thurs', 'Fri']

Vineek(TIMESLOTS = TIMESLOTS,
       DAYS = DAYS)
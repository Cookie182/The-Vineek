# The Vineek
 Simple, readable, epic 'algorithm' that's designed to create timetables for all courses and their respective semesters while also allocating appropriate classrooms to them such that no clashes in any way will occur. The algorithm will also generate faculty and classroom specific timetables as well. Algorithm should work in a way such that any modifications done to the excel file should reflect in the timetable producted by the algorithm so that any user will have no need to ever touch the code for whatever reason ... as long as they follow the conventions 😎🐧:
 * All the timetables, along with the initial excel sheet to provide the necessary data will be on the desktop. If not, the file directory path for them will be shown
 * 'Lecture_hrs', 'Tut_hrs', 'Lab_hrs' represents the amount of lectures of each respective type of lecture there is for a particular subject.
 * If a subject does not have any lab or tutorial lectures, the 'TA' column for that subject should be left as blank
 * Continuing from the previous point, put a 0 respective columns of 'Lecture_hrs', 'Lab_hrs', 'Tut_hrs' if the subject doesn't have any of those respective type of classes
 * 'Assigned_Room' and 'Assigned_Lab' columns are for pre-allocating specific rooms for a subject, either for a lecture or for a lab/tutorial session respectively
 * Please do not use the same values for 'Course_id' and 'Track_Core' columns
 * Please do make sure all variables in the 'Course_id' feature are unique as they play a integral role for track cores
 * Speaking of track cores, if a subject is common in a track core, allocate them in the same slot, assuming they also have the same teacher.
 * Try to keep the sample size small. To do this, try to only specific specific rooms to use for a specific course as the algorithm may panic and/or freeze
 * All the final outputs are in an excel format for further modifications and/or to cross reference to make more subjective changes
 * If you do wish to change the timeslots for all lectures, or even timeslots for labs as the amount of labs (assuming there are labs for the subject, are in multiples of 2) will be held consecutively, they can be changed within the code itself but in a very easy format just by editing the timeslots respectively near the end of the .py script.
 * Be gentle, she's a shy kind-hearted soul
 
### As an example of how data should be filled, look at the *Time-table.xlsx* file above

### Credits:
* Algorithm created by: Ashwin Jawalikar (me).
* Data gathering and logistics done by: [Gurvinder Kaur](https://github.com/gurvinder08)

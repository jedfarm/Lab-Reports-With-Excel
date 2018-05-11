# Physics Lab Reports With Excel

## Files

- Roster.csv (A text file containing the student roster, sample provided)

- SampleLabReport.xlsx (Template for the lab report, only the first and the last worksheets are required)

- CreateRoster.bas (this file needs Roster.csv to be located in the same directory as SampleLabReport, where this code is executing. It generates a drop down menu with the students' names in the first worksheet).

- CreateFeedbackSheet.bas (this macro requires a worksheet called RUBRIC to exist in SampleLabReport where this code is executing. It generates a new worksheet called Feedback where the instructor will enter the student's grades)

- collectLabGrades.py (This Python script scans a given folder for graded lab reports, collects all the grades and creates a .csv file with them in such a way that uploading it to CANVAS is straightforward. It needs the actual roster to be a in a known path)

## Tips for a better performance:

- It is wise to import both CreateRoster.bas and CreateFeedbackSheet.bas to PERSONAL.XLSB, so this code could be called to run on any Excel file.

- It would be a good practice for the instructor to make the RUBRIC worksheet password protected once finished customizing it. The students will have access, but won't be able to modify it in any way.

 - It does not matter which values the drop down menu had initially, when CreateRoster.bas is run, it overwrites the existing values.

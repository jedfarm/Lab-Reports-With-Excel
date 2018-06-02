# Science Lab Reports With Excel

This repository contains files that allow instructors to implement paperless Science Lab Reports. It was initially thought for Physics Labs (College level), but in fact, it could be used in a broader set of laboratories. Also, it was made to match the student roster provided by CANVAS, as it is the one we are currently using. 

The template "SampleLabReport.xslx" contains an actual lab we perform, as an example of the structure. Only the first and the last sheets in that file are relevant. In the first one, just the rows containing the team info have to be kept if the user intends to take advantage of the automation capabilities we are going to explain in the next section.

## How to use these files

- Download all the files from this repository to a local machine by going to the top right of the page, then click on the green button ** clone or download **,  then select Download ZIP.

- Import CreateRoster.bas and CreateFeedbackSheet.bas into Excel. 
If you don't know how to do it, [here](https://github.com/jedfarm/zipgrade/blob/master/README.md), I explain that for a different file. The procedure is the same. 

- Create your own lab reports templates, using the SampleLabReport.xlsx provided.  That includes customizing the RUBRIC tab, where the number of indicators could be changed at will. 

- Download the rosters from CANVAS as .csv files and delete all the columns except for those shown in the Roster.csv sample file.

- Create a separate folder for each lab group with a file named Roster.csv in it (generated in the previous step).

- Once the students have submitted their lab reports (it means that some rows of TEAM MEMBERS are going to be populated), run the macro CreateFeedbackSheet.  A new sheet called Feedback will appear on a given lab report file. There, as the instructor is typing grades in the feedback sheet, the lab grade will be updating itself in real time.

- After finish grading all the labs in a given folder; it is time to run collectLabGrades.py (for that Python must be installed in your system. Check for instance how to install Anaconda). It creates a file .cvs file in the same folder with all the grades, ready to be uploaded to Canvas.

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
 
 - The Sub CreateRosterForAllExcelFilesInFolder opens a dialog to select a folder and allows to run Create Roster at once for all the existing Excel files in that folder.

## Known limitations:
- The maximum number of students allowed per team is 4.

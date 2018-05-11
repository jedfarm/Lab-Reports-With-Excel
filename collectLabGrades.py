# -*- coding: utf-8 -*-
"""
Created on Sun Oct 22 22:17:00 2017
Scan all the Excel Lab reports in a folder and creates a .csv file in the same
folder with the grades ready to be uploaded to CANVAS
@author: jedfarm

"""

import os
import glob
import pandas as pd
import numpy as np

# This is the path to the folder containing all the labs reports that are already graded
# It must be changed every time.
mypath = 'C:/Users/jfand/Downloads/Calorimetry'

# This is where the roster for that group is. Usually that path doesn't change
# If you have multiple groups, then yes.
filepath_roster =  "C:/Users/jfand/Downloads/Rosters/SP18PHY2048L07981Roster.csv"




os.chdir(mypath)
FileList = glob.glob('*.xlsx')

pathList = []
for file in FileList:
    pathList.append(mypath + '/' + file)

xls_file = pd.ExcelFile(pathList[0])
df = xls_file.parse('Intro', parse_cols = [0])
labname = df.columns[0] + " Lab Report"
filename_out = df.columns[0]+"_grades_to_canvas.csv"
grades = pd.DataFrame()
for path in pathList:
    xls_file = pd.ExcelFile(path)
    df = xls_file.parse('Feedback', parse_cols = [1,2])
    df = df[['TEAM MEMBERS', 'GRADE']]
    stop = np.where(df['TEAM MEMBERS']== "INDICATORS")[0][0]
    grades = grades.append(df.iloc[:stop].dropna(subset=["GRADE"]))

grades.rename(columns={'GRADE': labname}, inplace=True)

# Depends on the given course
max_grade = 80




roster = pd.read_csv(filepath_roster)

gradesToCanvas = roster.merge(grades, left_on='Student', right_on='TEAM MEMBERS', how='outer')
del gradesToCanvas['TEAM MEMBERS']
gradesToCanvas.at[0, labname] = int(max_grade)
gradesToCanvas.to_csv(filename_out, encoding='utf-8', index=False)

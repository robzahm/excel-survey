import sys
import os
import shutil
import pandas as pd

print("Python Version: " + sys.version)

# Find directory with files
# Assumes the root directory contains the master template, and a subdirectory named "data"
# that contains the individual Excel files
rootDir = "/Users/zahm/surveys/"
surveyResultsMasterPath = rootDir + "survey_results_master.xlsx"
surveyResultsPath = rootDir + "survey_results.xlsx"

# Copy the survey results master into a new survey results file
shutil.copyfile(surveyResultsMasterPath, surveyResultsPath)

print ("Root Directory: " + rootDir)

dataDir = rootDir + "data/"
dataFiles = os.listdir(dataDir)

surveyResults = pd.ExcelFile(surveyResultsPath)
resultDF = surveyResults.parse("Daily Issue Survey")

# Load up a data file
dataFilePath = dataDir + "Copy of Survey - mistakes.xlsx"
dataFile = pd.ExcelFile(dataFilePath)
df = dataFile.parse("Daily Issue Survey")

print(df.iloc[5,0])

resultDF.iloc[5,0] = 2

print(resultDF.iloc[5,0])

writer = pd.ExcelWriter(surveyResultsPath, engine='xlsxwriter')
resultDF.to_excel(writer, "Daily Issue Survey")
writer.save()
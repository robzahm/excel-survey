import sys
import os
import shutil
import pandas as pd

# Function to add an amount to an existing or null value
def addValue(originalVal, amountToAdd):
    if not isinstance(originalVal, int):
        return amountToAdd
    else:
        return originalVal + amountToAdd


# Constants
NUM_DAILY_ENTRIES_ROW = 21
NUM_DAYS = 20
NUM_ISSUES = 19
REPORTING_SHEET_NAME = "Daily Issue Survey"
ROOT_DIRECTORY = "/Users/zahm/surveys/"

print("Python Version: " + sys.version)
print ("Looking for surveys in: " + ROOT_DIRECTORY)

# Copy the survey results master into a new survey results file,
# and load it into a dataframe
surveyResultsMasterPath = ROOT_DIRECTORY + "survey_results_master.xlsx"
surveyResultsPath = ROOT_DIRECTORY + "survey_results.xlsx"
shutil.copyfile(surveyResultsMasterPath, surveyResultsPath)
surveyResults = pd.ExcelFile(surveyResultsPath)
resultDF = surveyResults.parse(REPORTING_SHEET_NAME)


# Iterate over the data files
dataDir = ROOT_DIRECTORY + "data/"
dataFiles = os.listdir(dataDir)
for filename in dataFiles:
    print ("Processing Survey: " + filename)
    # Load the Excel file and survey entry sheet into a dataframe
    dataFile = pd.ExcelFile(dataDir + filename)
    df = dataFile.parse(REPORTING_SHEET_NAME)

    # Iterate over each column
    for colIndex in range(0,NUM_DAYS):
        valueCheckedThatDay = False

        # Iterate over each row except for row 0 which will have the date
        for rowIndex in range(1,NUM_ISSUES + 1):

            # Get the value of the dataframe for that cell
            dataVal = df.iloc[rowIndex, colIndex]

            # If the value is a number, we want to add it to our results
            if isinstance(dataVal, int):
                # Indicate that values were tracked this day, and update our results dataframe
                valueCheckedThatDay = True
                resultDF.iloc[rowIndex, colIndex] = addValue(resultDF.iloc[rowIndex, colIndex], dataVal)
        
        # If values were checked that day, increment the value in the result dataframe
        if valueCheckedThatDay:
            resultDF.iloc[NUM_DAILY_ENTRIES_ROW, colIndex] = addValue(resultDF.iloc[NUM_DAILY_ENTRIES_ROW, colIndex], 1)


# Write the output back to Excel
writer = pd.ExcelWriter(surveyResultsPath, engine='xlsxwriter')
resultDF.to_excel(writer, REPORTING_SHEET_NAME)
writer.save()

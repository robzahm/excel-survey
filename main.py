import sys
import os
import shutil
import pandas as pd

# Function to add an amount to an existing or null value
def add_value(originalVal, amountToAdd):
    if not isinstance(originalVal, int):
        return amountToAdd
    else:
        return originalVal + amountToAdd

# Function to exclude any opened Excel files
def list_non_hidden_files(path):
    for file in os.listdir(path):
        if not file.startswith("~") and file.endswith(".xlsx"):
            yield file

# Constants
NUM_DAILY_ENTRIES_ROW = 21
NUM_DAYS = 20
NUM_ISSUES = 19
REPORTING_SHEET_NAME = "Daily Issue Survey"
INSTRUCTIONS_SHEET_NAME = "Instructions"
ROOT_DIRECTORY = "/Users/zahm/surveys/"

print("Python Version: " + sys.version)
print ("Looking for surveys in: " + ROOT_DIRECTORY)
print()

# Copy the survey results master into a new survey results file,
# and load it into a dataframe
surveyResultsMasterPath = ROOT_DIRECTORY + "survey_results_master.xlsx"
surveyResultsPath = ROOT_DIRECTORY + "survey_results.xlsx"
shutil.copyfile(surveyResultsMasterPath, surveyResultsPath)
surveyResults = pd.ExcelFile(surveyResultsPath)
resultDF = surveyResults.parse(REPORTING_SHEET_NAME)


# Iterate over the data files
dataDir = ROOT_DIRECTORY + "data/"
dataFiles = list_non_hidden_files(dataDir)
for filename in dataFiles:
    print ("Processing Survey: " + filename)
    # Load the Excel file and survey entry sheet into a dataframe
    dataFile = pd.ExcelFile(dataDir + filename)
    df = dataFile.parse(sheet_name=REPORTING_SHEET_NAME)

    instructionsDF = dataFile.parse(sheet_name=INSTRUCTIONS_SHEET_NAME)
    hearingOffice = str(instructionsDF.iloc[0,3])
    region = str(instructionsDF.iloc[1,3])
    NHCLocation = str(instructionsDF.iloc[5,3])

    print ("Hearing Office: " + hearingOffice + ", Region: " + region + ", NHC Location: " + NHCLocation)
    print()

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
                resultDF.iloc[rowIndex, colIndex] = add_value(resultDF.iloc[rowIndex, colIndex], dataVal)
        
        # If values were checked that day, increment the value in the result dataframe
        if valueCheckedThatDay:
            resultDF.iloc[NUM_DAILY_ENTRIES_ROW, colIndex] = add_value(resultDF.iloc[NUM_DAILY_ENTRIES_ROW, colIndex], 1)

# Write the output back to Excel
writer = pd.ExcelWriter(surveyResultsPath, engine='xlsxwriter')
resultDF.to_excel(excel_writer=writer, sheet_name=REPORTING_SHEET_NAME)
writer.save()

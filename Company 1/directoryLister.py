import os
import pandas as pd
from openpyxl import load_workbook, Workbook


recipeList = []

# walk through directory / subdirectories and determine what recipes are here

for dirpath, dirnames, filenames in os.walk("."):
    for filename in [f for f in filenames if f.endswith(".rep")]: # or f.endswith(".txt") or f.endswith(".ini")]:
        file, extension = filename.split(".")
        if file not in recipeList:
            if "_" not in file:
                recipeList.append(file)

# load excel file

wb = load_workbook("DirectoryList.xlsx")

writer = pd.ExcelWriter("DirectoryList.xlsx", engine='openpyxl')

writer.book = wb
writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)

# write to excel file

# version 1 -----------------------------------------------------------------------------------------------
# create list of zeroes on righthand side which will be updated to 1's as needed later


# zeroes = []
#
# for i in range(len(recipeList)):
#     zeroes.append(0)
#
# recipeFrame = pd.DataFrame(list(zip(recipeList, zeroes)), columns=['Recipes', 'Include?'])
# recipeFrame.to_excel(writer, sheet_name='Sheet1', header=True, index=False, startcol=0, startrow=0)
# ---------------------------------------------------------------------------------------------------------


# version 2 -----------------------------------------------------------------------------------------------
recipeFrame = pd.DataFrame(recipeList, columns=['Recipes'])
i = 0

infoFrame = pd.DataFrame(["In the following pages, change the 0 in the upper left corner to a 1 if you want to run comparisons on that file."],columns= ['Info'])
infoFrame.to_excel(writer, sheet_name='Sheet1', header=None, index=False, startcol=0, startrow=0)

while i < len(list(recipeFrame.Recipes)):
    recipe = recipeFrame.loc[i].at["Recipes"]
    file = open(recipe + ".rep", "r")
    recipeParams = []
    criticals = []
    for line in file:
        parameter, value = line.split("=")
        recipeParams.append(parameter)
    for k in range(len(recipeParams)):
        criticals.append(1)
    paramFrame = pd.DataFrame(list(zip(criticals, recipeParams)), columns= ['Critical Keys', 'Parameters'])
    paramFrame.to_excel(writer, sheet_name=recipe, header=True, index=False, startcol=0, startrow=3)
    toggleFrame = pd.DataFrame([0],columns=['Toggle'])
    toggleFrame.to_excel(writer, sheet_name=recipe, header=None, index=False, startcol=0, startrow=0)
    fileFrame = pd.DataFrame([recipe],columns=['recipe'])
    fileFrame.to_excel(writer, sheet_name=recipe, header=None, index=False, startcol=1, startrow=0)
    i += 1
# ---------------------------------------------------------------------------------------------------------

# remember to save

wb.save("DirectoryList.xlsx")
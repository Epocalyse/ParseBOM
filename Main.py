import os
from Convert import Convert
from ExtractBOM import ExtractBOM

# Main file to run everything

print("SOP or BOM?")
fileType = input()
loc = ''

extract_BOM = ExtractBOM()
if fileType == "SOP":
    loc = "SOPWordCopies"
    for file in os.listdir("SOPWordCopies"):
        extract_BOM.convertDocx(file, loc)
        Convert.sopDocToCSV(file, loc)
elif fileType == "BOM":
    loc = "BOMWordCopies"
    for file in os.listdir("BOMWordCopies"):
        extract_BOM.convertDocx(file, loc)
        Convert.bomDocToCSV(file, loc)
else:
    print("Not a valid file type.")
    exit()
